#!/usr/bin/env node
import assert from 'node:assert/strict';
import {generate, verifySeal, prepareCopy, inspectNative, sealCandidate} from './pack.mjs';
const [command, ...args] = process.argv.slice(2), options = {};
assert.equal(args.length % 2, 0, 'Expected --name value pairs');
for (let index = 0; index < args.length; index += 2) {
  assert.match(args[index], /^--(?:consumer|output|suite|route|input|mode|candidate|source|revision|tarball)$/);
  const name = args[index].slice(2); assert.ok(!Object.hasOwn(options, name), `Duplicate --${name}`); options[name] = args[index + 1];
}
const required = {generate: ['consumer', 'output'], verify: ['suite'], copy: ['suite', 'route', 'output'], inspect: ['suite', 'consumer', 'input', 'route', 'output'], candidate: ['source', 'revision', 'tarball', 'consumer', 'output']};
assert.ok(Object.hasOwn(required, command), 'Commands: candidate, generate, verify, copy, inspect (docs/keynote-qualification.md)');
const optional = command === 'generate' ? ['mode', 'candidate'] : [];
assert.ok(required[command].every(key => Object.hasOwn(options, key)) && Object.keys(options).every(key => required[command].includes(key) || optional.includes(key)), 'Missing or unexpected options');
if (command === 'candidate') {const receipt = await sealCandidate(options); console.log(`Sealed local candidate ${receipt.source.revision}; no registry claim. ${options.output}`);}
if (command === 'generate') {const suite = await generate(options); console.log(`Prepared ${suite.kind}: ${suite.fixture ?? 'baseline-16'} / ${suite.mode ?? 'baseline'}; native acceptance pending. ${options.output}`);}
if (command === 'verify') {const {suite} = await verifySeal(options.suite); console.log(`Verified ${suite.files.length} sealed files; no native acceptance claim.`);}
if (command === 'copy') console.log(`Fresh editable attempt: ${await prepareCopy(options)}`);
if (command === 'inspect') {const report = await inspectNative(options); console.log(`Static native-export inspection: ${report.failures.length} findings. ${options.output}`); if (report.failures.length) process.exitCode = 1;}
