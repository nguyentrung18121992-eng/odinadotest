#!/usr/bin/env node
'use strict';

/**
 * Loads all serverless + bot modules without starting HTTP. Exits 0 on success.
 */
require('../services/botApp');
require('../api/messages');
console.log('verify: ok');
