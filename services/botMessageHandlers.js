'use strict';

const ado = require('./adoBoardsClient');

const defaultWorkItemType = process.env.ADO_DEFAULT_WORK_ITEM_TYPE || 'Task';

/**
 * After a command prefix, expect "<id> <rest...>".
 * @param {string} remainder
 * @returns {{ id: number, rest: string } | null}
 */
function parseIdAndRest(remainder) {
  const m = remainder.trim().match(/^(\d+)\s+([\s\S]+)$/);
  if (!m) {
    return null;
  }
  return { id: Number(m[1], 10), rest: m[2].trim() };
}

/**
 * @param {string} remainder
 * @returns {number | null}
 */
function parseIdOnly(remainder) {
  const m = remainder.trim().match(/^(\d+)$/);
  if (!m) {
    return null;
  }
  return Number(m[1], 10);
}

function helpText() {
  return [
    '**Commands**',
    '- `create task <title>` — new work item',
    '- `title <id> <new title>` — rename',
    '- `description <id> <text>` — set description',
    '- `assign <id> <email or name>` — set assignee',
    '- `unassign <id>` — clear assignee',
    '- `state <id> <state>` — change state (e.g. Active, Closed)',
    '- `update <id> <state>` — same as `state`',
    '- `search <keywords>` / `find <keywords>` — title & description (WIQL `CONTAINS WORDS`; `ADO_SEARCH_TOP` optional, max 50)',
    '- `search <id>` — when `<id>` is digits only, open that work item in the project',
    '- `delete <id>` / `delete task <id>` — delete work item',
  ].join('\n');
}

const MAX_SEARCH_TITLE_CHARS = 120;

/**
 * @param {import('azure-devops-node-api/interfaces/WorkItemTrackingInterfaces').WorkItem} wi
 * @param {string} refName
 */
function fieldDisplay(wi, refName) {
  const f = wi.fields?.[refName];
  if (f == null) {
    return '';
  }
  if (typeof f === 'object' && f.displayName != null) {
    return String(f.displayName);
  }
  return String(f);
}

/**
 * @param {string} s
 * @param {number} max
 */
function truncate(s, max) {
  if (s.length <= max) {
    return s;
  }
  return `${s.slice(0, max - 1)}…`;
}

/** @param {import('botbuilder').TurnContext} context */
async function handleCreateTask(context, text) {
  const m = text.match(/^create task\s+/i);
  if (!m) {
    return false;
  }
  const title = text.slice(m[0].length).trim();
  if (!title) {
    await context.sendActivity('Usage: `create task <title>`');
    return true;
  }
  const wi = await ado.createWorkItem({
    workItemType: defaultWorkItemType,
    fields: {
      title,
      'System.Description': 'Created from Microsoft Teams bot.',
    },
  });
  const url = ado.workItemWebUrl(wi);
  await context.sendActivity(
    `Created work item **#${wi.id}**${url ? `\n${url}` : ''}`
  );
  return true;
}

/** @param {import('botbuilder').TurnContext} context */
async function handleSearch(context, text, lower) {
  const mSearch = text.match(/^search\s+/i);
  const mFind = text.match(/^find\s+/i);
  const m = mSearch || mFind;
  if (!m) {
    return false;
  }
  const q = text.slice(m[0].length).trim();
  if (!q) {
    await context.sendActivity(
      'Usage: `search <keywords>` or `find <keywords>` — use `description <id> <text>` to read or set full description.'
    );
    return true;
  }
  try {
    const items = await ado.searchWorkItems({ text: q });
    if (items.length === 0) {
      await context.sendActivity(`No work items matched **${q}** in this project.`);
      return true;
    }
    const lines = items.map((wi) => {
      const id = wi.id;
      const wtype = fieldDisplay(wi, 'System.WorkItemType') || 'Work item';
      const state = fieldDisplay(wi, 'System.State') || '—';
      const title = truncate(
        fieldDisplay(wi, 'System.Title') || '(no title)',
        MAX_SEARCH_TITLE_CHARS
      );
      const assignee = fieldDisplay(wi, 'System.AssignedTo');
      const assignBit = assignee ? ` · ${assignee}` : '';
      const url = ado.workItemWebUrl(wi);
      const link = url ? `\n${url}` : '';
      return `• **#${id}** [${wtype}] *${state}* — ${title}${assignBit}${link}`;
    });
    const header = `**${items.length}** result(s) for \`${q}\``;
    await context.sendActivity([header, ...lines].join('\n'));
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    await context.sendActivity(`Search failed: ${msg}`);
  }
  return true;
}

/** @param {import('botbuilder').TurnContext} context */
async function handleDeleteTask(context, text, lower) {
  let id = null;
  if (lower.startsWith('delete task')) {
    const dm = text.match(/^delete\s+task\s+(\d+)\s*$/i);
    if (!dm) {
      await context.sendActivity('Usage: `delete task <id>`');
      return true;
    }
    id = Number(dm[1], 10);
  } else if (lower.startsWith('delete ')) {
    const dm = text.match(/^delete\s+(\d+)\s*$/i);
    if (!dm) {
      await context.sendActivity('Usage: `delete <id>`');
      return true;
    }
    id = Number(dm[1], 10);
  }
  if (id == null) {
    return false;
  }
  await ado.deleteWorkItem({ id, destroy: false });
  await context.sendActivity(`Deleted work item **#${id}** (sent to Recycle Bin).`);
  return true;
}

/** @param {import('botbuilder').TurnContext} context */
async function handleTitle(context, text, lower) {
  const m = text.match(/^title\s+/i);
  if (!m) {
    return false;
  }
  const parsed = parseIdAndRest(text.slice(m[0].length));
  if (!parsed) {
    await context.sendActivity('Usage: `title <id> <new title>`');
    return true;
  }
  const wi = await ado.updateWorkItem({
    id: parsed.id,
    fields: { 'System.Title': parsed.rest },
  });
  const url = ado.workItemWebUrl(wi);
  await context.sendActivity(
    `Updated title on **#${wi.id}**${url ? `\n${url}` : ''}`
  );
  return true;
}

/** @param {import('botbuilder').TurnContext} context */
async function handleDescription(context, text, lower) {
  const m = text.match(/^description\s+/i);
  if (!m) {
    return false;
  }
  const parsed = parseIdAndRest(text.slice(m[0].length));
  if (!parsed) {
    await context.sendActivity('Usage: `description <id> <text>`');
    return true;
  }
  const wi = await ado.updateWorkItem({
    id: parsed.id,
    fields: { 'System.Description': parsed.rest },
  });
  const url = ado.workItemWebUrl(wi);
  await context.sendActivity(
    `Updated description on **#${wi.id}**${url ? `\n${url}` : ''}`
  );
  return true;
}

/** @param {import('botbuilder').TurnContext} context */
async function handleAssign(context, text, lower) {
  const m = text.match(/^assign\s+/i);
  if (!m) {
    return false;
  }
  const parsed = parseIdAndRest(text.slice(m[0].length));
  if (!parsed) {
    await context.sendActivity('Usage: `assign <id> <email or display name>`');
    return true;
  }
  const wi = await ado.updateWorkItem({
    id: parsed.id,
    fields: { 'System.AssignedTo': parsed.rest },
  });
  const url = ado.workItemWebUrl(wi);
  await context.sendActivity(
    `Assigned **#${wi.id}** to \`${parsed.rest}\`${url ? `\n${url}` : ''}`
  );
  return true;
}

/** @param {import('botbuilder').TurnContext} context */
async function handleUnassign(context, text, lower) {
  const m = text.match(/^unassign\s+/i);
  if (!m) {
    return false;
  }
  const id = parseIdOnly(text.slice(m[0].length));
  if (id == null || Number.isNaN(id)) {
    await context.sendActivity('Usage: `unassign <id>`');
    return true;
  }
  const wi = await ado.updateWorkItem({
    id,
    fields: { 'System.AssignedTo': '' },
  });
  const url = ado.workItemWebUrl(wi);
  await context.sendActivity(
    `Unassigned **#${wi.id}**${url ? `\n${url}` : ''}`
  );
  return true;
}

/** @param {import('botbuilder').TurnContext} context */
async function handleState(context, text, lower) {
  const mState = text.match(/^state\s+/i);
  const mUpdate = text.match(/^update\s+/i);
  const m = mState || mUpdate;
  if (!m) {
    return false;
  }
  const parsed = parseIdAndRest(text.slice(m[0].length));
  if (!parsed) {
    await context.sendActivity(
      'Usage: `state <id> <state>` — example: `state 123 Active`'
    );
    return true;
  }
  const wi = await ado.updateWorkItem({
    id: parsed.id,
    fields: { 'System.State': parsed.rest },
  });
  const url = ado.workItemWebUrl(wi);
  await context.sendActivity(
    `**#${wi.id}** → state \`${parsed.rest}\`${url ? `\n${url}` : ''}`
  );
  return true;
}

const handlers = [
  handleCreateTask,
  handleSearch,
  handleDeleteTask,
  handleTitle,
  handleDescription,
  handleAssign,
  handleUnassign,
  handleState,
];

/**
 * @param {import('botbuilder').TurnContext} context
 */
async function handleTeamsMessage(context) {
  const text = (context.activity.text || '').trim();
  if (!text) {
    await context.sendActivity(helpText());
    return;
  }
  const lower = text.toLowerCase();

  for (const fn of handlers) {
    // eslint-disable-next-line no-await-in-loop
    const handled = await fn(context, text, lower);
    if (handled) {
      return;
    }
  }

  await context.sendActivity(helpText());
}

module.exports = {
  handleTeamsMessage,
  helpText,
};
