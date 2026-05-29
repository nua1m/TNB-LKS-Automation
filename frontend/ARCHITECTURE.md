Frontend rebuild plan

Stack
- Vite
- React
- TypeScript
- Tailwind CSS
- shadcn/ui component patterns

Why this stack
- The current Qt-widget shell cannot reach the UI quality target.
- The desktop backend is still Python. That part already works.
- The frontend should become a proper component system, not hand-made page styling.

Desktop contract
- The Python host remains responsible for file dialogs, updater calls, and Excel workflows.
- The frontend only renders state and sends user actions.
- The frontend expects these bridge methods:
  - `getInitialState()`
  - `pickFile(kind)`
  - `pickLksFiles()`
  - `pickDirectory()`
  - `openPath(target)`
  - `checkUpdates(module)`
  - `startLks(payloadJson)`
  - `startPayslip(payloadJson)`
  - `respondAppendConfirmation(answer)`
- The frontend listens to one event stream:
  - `eventEmitted(payloadJson)`

Information architecture
- Left sidebar navigation only
- No logo icon
- No status pill in the header
- One small info action in the header
- Two workspaces:
  - LKS Automation
  - Payslip Generator

Design rules
- Neutral business UI
- Consistent spacing scale
- Consistent radius scale
- No hero-marketing styling
- No decorative gradients except very restrained surfaces
- Dense enough for desktop work, not consumer-app loose spacing

Build output
- Vite builds into `../web_ui`
- The Python shell continues loading `web_ui/index.html`
- That means the new frontend replaces the old static prototype at build time
