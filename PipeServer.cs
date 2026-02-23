// PipeServer.cs — OBSOLETE
//
// The original stdio-redirect approach (MCP server running elevated with
// stdin/stdout redirected over named pipes) has been replaced by the
// proxy architecture:
//
//   --mcp-connect: Non-elevated MCP stdio server (in Program.cs)
//                  proxies UIA calls via UiaProxyClient -> named pipe
//
//   --server:      Elevated UiaProxyServer (UiaProxy.cs)
//                  executes UIA commands and returns JSON results
//
// Auto-launch and mutex logic now lives directly in Program.cs.
// This file is kept as documentation only.
