#!/usr/bin/env node

import { defineCommand, runMain } from "citty";
import { assignments } from "./commands/assignments.js";
import { debug } from "./commands/debug.js";
import { example } from "./commands/example.js";

const main = defineCommand({
  meta: {
    name: "quadro",
    version: "0.1.0",
    description: "Quadro CLI — Excel workbook generation tools",
  },
  subCommands: {
    example,
    assignments,
    debug,
  },
});

runMain(main);
