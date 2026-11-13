# Markdown Parser Fallback Heuristic Disabling & Tool Stream Hardening

**Status**: Fixed (2026-10-24)  
**Area**: Agent tool stream parsing · Markdown heuristic fallbacks · Schema-compliant execution

## Problem

1. When streaming agent tool calls or multi-turn outputs containing code blocks, shell commands, or structured parameters, aggressive fallback heuristics in Markdown parsers attempted to "auto-repair" incomplete syntax.
2. These heuristics frequently misinterpreted raw text, unclosed backticks, or nested lists inside Markdown code blocks as broken syntax, improperly escaping characters or mangling JSON payloads before dispatch.
3. In conversational agent execution loops, this caused spurious syntax repair passes, corrupting structured tool envelopes and causing execution failures or schema rejections.

## Root Cause (Why)

```text
FAILURE PATH (Aggressive Fallback Enabled)
  Raw Tool Envelope / Streamed Markdown
         |
         v
  Parser detects trailing unclosed token or formatted snippet
         |
         v
  Aggressive Heuristic Triggers:
    Attempts synthetic closure / auto-wrapping
         |
         v
  Payload mutated -> JSON parsing error / Command schema mismatch
         |
         v
  Execution FAILS / Spurious Error Injected

SUCCESS PATH (Fallback Heuristics Disabled)
  Raw Tool Envelope / Streamed Markdown
         |
         v
  Strict Deterministic Lexer / Parser
         |
         v
  Preserves exact character sequence & raw string boundaries
         |
         v
  Zero Mutation -> Uncompromised JSON & Schema Validation
         |
         v
  Tool Execution SUCCEEDS cleanly
```

1. **Speculative Heuristics**: Fallback heuristics assumed non-standard Markdown should be aggressively coerced into valid HTML ASTs, rather than preserving raw code literals for tool runners.
2. **Double-Escaping Traps**: Auto-repair routines inserted backslashes or entity escapes into stringified JSON parameters (such as `{\"type\":\"command\"}`), causing subsequent JSON deserializers to reject the structure.

## Fix (How)

1. **Disabled Parser Fallback Heuristic**: Strict deterministic lexing is enforced without speculative syntax completion passes or lossy coercion.
2. **Direct Buffer Ingestion**: Tool calls and payload blocks preserve strict character invariants without intermediate sanitizers stripping quotes, brackets, or backticks.
3. **Deterministic Boundary Checking**: Unclosed blocks or partial frames are flagged directly as incomplete rather than synthetically rewritten.

## Files & Functions

| File | Functions / Role |
| :--- | :--- |
| [`ExcelCommandExecutor.ts`](../../src/websocket/service/ExcelCommandExecutor.ts#L161-L295) | Deterministic envelope serialization and error boundary checking |
| [`BackgroundService.ts`](../../src/websocket/service/BackgroundService.ts#L190-L232) | Clean command dispatch without mutation or lossy JSON transformations |
| [`MessageHandler.ts`](../../src/websocket/client/MessageHandler.ts#L45-L85) | Strict message frame validation and routing |
