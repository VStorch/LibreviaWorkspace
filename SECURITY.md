# Security policy

**English** · [Português (Brasil)](SECURITY.pt-BR.md)

Librevia opens files that come from other people. A `.docx` or an `.xlsx` may have been crafted
on purpose to crash the application or to try to escape the place where it runs. That is why
security reports are taken seriously and handled privately.

## Supported versions

Only the **latest version**. The project does not yet maintain old versions in parallel; the fix
ships in a new version.

## How to report

**Do not open a public issue.** A flaw described in public is visible to whoever wants to exploit
it before there is a fix.

Use GitHub's **private vulnerability reporting**:

1. open the repository's **Security** tab;
2. click **Report a vulnerability**;
3. describe the problem.

It helps a lot to include:

- the Librevia version and the operating system;
- the steps to reproduce;
- the file that triggers the problem, if there is one, or a description of how to build it;
- what happens, and what you expected to happen.

## What happens next

- You receive a reply confirming receipt.
- The problem is reproduced and assessed.
- The fix ships in a new version, and the security advisory is published with it, with your name
  in the credits, if you wish.

The project is maintained by a single person, so the timelines are good-faith, not contractual.

## In scope

The application was designed around a few promises. Breaking any of them is a security flaw:

| Promise | Example of a flaw |
| --- | --- |
| **A file does not bring the application down** | a `.docx` or `.xlsx` that crashes the format service, exhausts memory or opens a ZIP that expands without end |
| **The interface runs isolated** | a document that manages to run code outside the renderer sandbox, or to reach Node.js |
| **A PDF stays in the viewer** | a PDF that, when opened, reaches the application's page or functions, opens a window, navigates the application window away or reads any file other than itself |
| **A link only leaves with confirmation** | a link in a file that opens the browser without the user seeing the address and accepting |
| **Writes only go where the user chose** | writing or reading a path that went through neither a system dialog nor the recent files list, bypassing the `PATH_NOT_AUTHORIZED` lock |
| **The application uses the network only for the AI assistant** | a document that makes the application fetch something from the internet on its own, on opening or on printing; any connection without an assistant configured, except the `localhost` query when the person clicks "Find assistants" |
| **The assistant sends only the document on the active tab or the active sheet, and only with permission** | text from another file, from another tab of the window, from another sheet tab, the file name or the properties going to the assistant; a send outside the computer without the confirmation made by the main process, which says the document on the active tab or the active sheet goes along as context; a conversation tool that changes the document without the person's click or that reads anything beyond the text sent with the message: the tools only read that text and propose |
| **Librevia does not give the AI agent access to the computer** | Claude Code running with any tool of its own or with another MCP server (it is called with `--tools ""` and `--strict-mcp-config`, and only receives the five document tools); Librevia offering an ACP agent files or a terminal, or accepting a permission request from it; the agent running outside an empty temporary folder; the agent's program coming from a path chosen by the interface, without going through a system dialog; the process or its children staying alive after stopping or after the window closes. An ACP agent's own tools (those of Gemini CLI, Codex, Claude Agent…) follow its configuration, not Librevia's: whoever wants to restrict them configures the agent |
| **The MCP server serves only its own user, and only proposes** | the server listening on a TCP port, or turned on without the person having ticked the box; the socket or its folder with permission for another user (the folder is 700, the socket 600); a connection served without the correct key, after the 2-second deadline, with the old key after the app reopens or the box is ticked again, or after it is turned off; more than 2 programs connected; a connected program reading or changing a tab other than the active one of the window in front, reading a file, path, file name, properties or PDF, saving, accepting or rejecting revisions, changing the document other than through a tracked change or a comment whose author ends in "(MCP)", or changing a spreadsheet cell without the person clicking Apply on the proposal; the key of a Claude Code message staying valid after the reply ends or reaching the live document |
| **The assistant's key does not leak** | the key going back to the interface, showing up in a log or stored unencrypted on disk |
| **The format service writes nothing** | the .NET sidecar writing to disk or opening a connection |
| **A PDF's properties only append** | writing a PDF's properties changes some byte of what was already in the file, or writes to a PDF other than the one opened |

## Out of scope

- Files over 20 MB, which the application already refuses before reading.
- Macros in a `.dotm`: they are neither run nor copied into the new document.
- Problems that require someone to already have control of the user's computer.
- Flaws in Electron, Chromium or .NET that already have a fix published by them. In that case,
  let us know that Librevia is on an outdated version; that is welcome, but as a regular issue.
