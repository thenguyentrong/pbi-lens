// MCP over stdio: stdout carries ONLY JSON-RPC. Any stray console.log from a
// dependency (express, MSAL, playwright) would corrupt the protocol, so all
// console output is rerouted to stderr. Import this module FIRST.
/* eslint-disable no-console */
console.log = console.error.bind(console);
console.info = console.error.bind(console);
console.warn = console.error.bind(console);
console.debug = console.error.bind(console);

export {};
