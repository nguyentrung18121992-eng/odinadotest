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
    '- `delete <id>` / `delete task <id>` — delete work item',
  ].join('\n');
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
