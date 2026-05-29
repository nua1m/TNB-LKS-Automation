import { type ReactNode, useEffect, useMemo, useState } from "react";

import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Separator } from "@/components/ui/separator";
import { cn } from "@/lib/utils";

type ModuleKey = "lks" | "payslip";

type PayslipRunResult = Extract<
  DesktopEvent,
  { type: "runCompleted"; module: "payslip" }
>["result"];

type DesktopEvent =
  | { type: "toast"; level?: "info" | "success" | "error"; message: string }
  | { type: "log"; module: ModuleKey; message: string }
  | { type: "runStarted"; module: ModuleKey }
  | { type: "status"; module: "lks"; value: string }
  | {
      type: "appendConfirmation";
      module: "lks";
      existingCount: number;
      newCount: number;
    }
  | {
      type: "runCompleted";
      module: "lks";
      result: {
        aborted: boolean;
        outputPath?: string;
        generatedInputPath?: string;
        summary?: Record<string, string | number>;
        trasByDate?: Record<string, number>;
      };
    }
  | {
      type: "runCompleted";
      module: "payslip";
      result: {
        outputDir: string;
        generatedXlsxCount: number;
        generatedPdfCount: number;
        warnings: string[];
        pdfFailures: string[];
        calculationWorkbookPath?: string;
        claimSummary?: {
          sourceFiles: number;
          totalRows: number;
          countedRows: number;
          skippedRows: number;
          fileSummaries: Array<{
            fileName: string;
            totalRows: number;
            countedRows: number;
            skippedRows: number;
          }>;
        } | null;
      };
    }
  | { type: "runFailed"; module: ModuleKey; message: string };

type DesktopInitialState = {
  version: string;
  supportEmail: string;
  supportPhone: string;
  defaults: {
    lksTemplate?: string;
    calcPath?: string;
    masterPath?: string;
    outputDir?: string;
  };
};

type DesktopBridge = {
  eventEmitted: {
    connect: (handler: (payload: string) => void) => void;
  };
  getInitialState: (callback: (value: string) => void) => void;
  pickFile: (kind: string, callback: (value: string) => void) => void;
  pickLksFiles: (callback: (value: string) => void) => void;
  pickDirectory: (callback: (value: string) => void) => void;
  openPath: (target: string) => void;
  checkUpdates: (module: string) => void;
  startLks: (payloadJson: string) => void;
  startPayslip: (payloadJson: string) => void;
  respondAppendConfirmation: (answer: boolean) => void;
};

declare global {
  interface Window {
    qt?: unknown;
    __tnbDispatchEvent?: (payload: string) => void;
    QWebChannel?: new (
      transport: unknown,
      callback: (channel: { objects: { bridge: DesktopBridge } }) => void,
    ) => void;
  }
}

type ToastState = {
  id: number;
  level: "info" | "success" | "error";
  message: string;
};

type LksState = {
  running: boolean;
  inputPath: string;
  outputPath: string;
  generatedInputPath: string;
  logLines: string[];
  summary: Record<string, string | number> | null;
  trasByDate: Record<string, number>;
  statusText: string;
};

type PayslipState = {
  running: boolean;
  calcPath: string;
  masterPath: string;
  outputDir: string;
  salaryMonth: string;
  paymentDate: string;
  lksPaths: string[];
  lastOutputDir: string;
  logLines: string[];
  summary: PayslipRunResult | null;
};

const moduleMeta: Record<
  ModuleKey,
  {
    label: string;
    title: string;
    description: string;
  }
> = {
  lks: {
    label: "Operations",
    title: "LKS Automation",
    description:
      "Prepare the LKS workbook, review duplicates and TRAS output, and open the generated files.",
  },
  payslip: {
    label: "Payroll",
    title: "Payslip Generator",
    description:
      "Generate salary workbooks and PDF payslips from the calculation workbook or from selected LKS CLAIM files.",
  },
};

const lksSummarySchema: Array<{ key: string; label: string }> = [
  { key: "Processed SOs", label: "Processed SOs" },
  { key: "Added to template", label: "Added" },
  { key: "Duplicate SOs skipped", label: "Duplicate SOs" },
  { key: "SOs with duplicates", label: "Duplicate groups" },
  { key: "Rows needing review", label: "Need review" },
  { key: "Rows skipped for TRAS", label: "TRAS skipped" },
  { key: "Missing OLD meter", label: "Missing old" },
  { key: "Missing CARD", label: "Missing card" },
  { key: "Missing NEW meter", label: "Missing new" },
  { key: "Execution time", label: "Execution time" },
];

function App() {
  const [bridge, setBridge] = useState<DesktopBridge | null>(null);
  const [activeModule, setActiveModule] = useState<ModuleKey>("lks");
  const [infoOpen, setInfoOpen] = useState(false);
  const [appendOpen, setAppendOpen] = useState(false);
  const [appendPrompt, setAppendPrompt] = useState({ existing: 0, incoming: 0 });
  const [version, setVersion] = useState("0.0.0");
  const [supportEmail, setSupportEmail] = useState("syahmi@nuaim.my");
  const [supportPhone, setSupportPhone] = useState("+60 18 2605 390");
  const [toasts, setToasts] = useState<ToastState[]>([]);

  const [lks, setLks] = useState<LksState>({
    running: false,
    inputPath: "",
    outputPath: "",
    generatedInputPath: "",
    logLines: ["No run started yet."],
    summary: null,
    trasByDate: {},
    statusText: "Ready",
  });

  const [payslip, setPayslip] = useState<PayslipState>({
    running: false,
    calcPath: "",
    masterPath: "",
    outputDir: "",
    salaryMonth: formatMonthLabel(new Date()),
    paymentDate: formatDateInput(new Date()),
    lksPaths: [],
    lastOutputDir: "",
    logLines: ["No generation started yet."],
    summary: null,
  });

  useEffect(() => {
    let cancelled = false;
    let attempts = 0;

    window.__tnbDispatchEvent = (payload: string) => {
      try {
        handleBridgeEvent(payload);
      } catch (error) {
        addToast("error", `UI event error: ${String(error)}`);
      }
    };

    const connectBridge = () => {
      if (cancelled) {
        return;
      }

      if (!window.qt || !window.QWebChannel) {
        attempts += 1;
        if (attempts > 40) {
          addToast("error", "Desktop bridge is not available. Open this UI from the desktop app.");
          return;
        }
        window.setTimeout(connectBridge, 100);
        return;
      }

      new window.QWebChannel(
        (window.qt as { webChannelTransport: unknown }).webChannelTransport,
        (channel) => {
          if (cancelled) {
            return;
          }

          const nextBridge = channel.objects.bridge;
          setBridge(nextBridge);
          nextBridge.eventEmitted.connect((payload) => handleBridgeEvent(payload));
          bridgeCall<string>(nextBridge, "getInitialState")
            .then((raw) => {
              const initial = JSON.parse(raw) as DesktopInitialState;
              setVersion(initial.version || "0.0.0");
              setSupportEmail(initial.supportEmail || "syahmi@nuaim.my");
              setSupportPhone(initial.supportPhone || "+60 18 2605 390");
              setPayslip((current) => ({
                ...current,
                calcPath: initial.defaults.calcPath || "",
                masterPath: initial.defaults.masterPath || "",
                outputDir: initial.defaults.outputDir || "",
              }));
            })
            .catch((error) => addToast("error", String(error)));
        },
      );
    };

    connectBridge();

    return () => {
      cancelled = true;
      delete window.__tnbDispatchEvent;
    };
  }, []);

  function handleBridgeEvent(rawPayload: string) {
    const payload = JSON.parse(rawPayload) as DesktopEvent;

    if (payload.type === "toast") {
      addToast(payload.level || "info", payload.message);
      return;
    }

    if (payload.type === "log" && payload.module === "lks") {
      setLks((current) => ({
        ...current,
        logLines: appendUniqueLog(current.logLines, "No run started yet.", payload.message),
      }));
      return;
    }

    if (payload.type === "log" && payload.module === "payslip") {
      setPayslip((current) => ({
        ...current,
        logLines: appendUniqueLog(current.logLines, "No generation started yet.", payload.message),
      }));
      return;
    }

    if (payload.type === "runStarted" && payload.module === "lks") {
      setLks((current) => ({
        ...current,
        running: true,
        outputPath: "",
        generatedInputPath: "",
        summary: null,
        trasByDate: {},
        statusText: "Running",
        logLines: ["Starting LKS automation..."],
      }));
      return;
    }

    if (payload.type === "runStarted" && payload.module === "payslip") {
      setPayslip((current) => ({
        ...current,
        running: true,
        lastOutputDir: "",
        summary: null,
        logLines: ["Starting payslip generation..."],
      }));
      return;
    }

    if (payload.type === "status") {
      setLks((current) => ({
        ...current,
        statusText: payload.value,
      }));
      return;
    }

    if (payload.type === "appendConfirmation") {
      setAppendPrompt({
        existing: payload.existingCount,
        incoming: payload.newCount,
      });
      setAppendOpen(true);
      return;
    }

    if (payload.type === "runCompleted" && payload.module === "lks") {
      setLks((current) => ({
        ...current,
        running: false,
        outputPath: payload.result.outputPath || "",
        generatedInputPath: payload.result.generatedInputPath || "",
        summary: payload.result.summary || null,
        trasByDate: payload.result.trasByDate || {},
        statusText: payload.result.aborted ? "Aborted" : "Completed",
        logLines: [
          ...trimPlaceholder(current.logLines, "No run started yet."),
          payload.result.aborted ? "Run aborted." : "Run completed.",
        ],
      }));
      addToast(payload.result.aborted ? "info" : "success", payload.result.aborted ? "LKS run was cancelled." : "LKS workbook generated.");
      return;
    }

    if (payload.type === "runCompleted" && payload.module === "payslip") {
      setPayslip((current) => ({
        ...current,
        running: false,
        lastOutputDir: payload.result.outputDir || "",
        summary: payload.result,
        logLines: [
          ...trimPlaceholder(current.logLines, "No generation started yet."),
          "Generation completed.",
        ],
      }));
      addToast("success", "Payslip generation completed.");
      return;
    }

    if (payload.type === "runFailed" && payload.module === "lks") {
      setLks((current) => ({
        ...current,
        running: false,
        statusText: "Failed",
        logLines: [...trimPlaceholder(current.logLines, "No run started yet."), `Run failed: ${payload.message}`],
      }));
      addToast("error", payload.message);
      return;
    }

    if (payload.type === "runFailed" && payload.module === "payslip") {
      setPayslip((current) => ({
        ...current,
        running: false,
        logLines: [
          ...trimPlaceholder(current.logLines, "No generation started yet."),
          `Generation failed: ${payload.message}`,
        ],
      }));
      addToast("error", payload.message);
      return;
    }
  }

  function addToast(level: ToastState["level"], message: string) {
    const id = Date.now() + Math.floor(Math.random() * 1000);
    setToasts((current) => [...current, { id, level, message }]);
    window.setTimeout(() => {
      setToasts((current) => current.filter((toast) => toast.id !== id));
    }, 4200);
  }

  async function browseFile(kind: "lksInput" | "calc" | "master") {
    if (!bridge) {
      addToast("error", "Desktop bridge not ready.");
      return;
    }
    try {
      const selected = await bridgeCall<string>(bridge, "pickFile", kind);
      if (!selected) {
        return;
      }
      if (kind === "lksInput") {
        setLks((current) => ({ ...current, inputPath: selected }));
      }
      if (kind === "calc") {
        setPayslip((current) => ({ ...current, calcPath: selected }));
      }
      if (kind === "master") {
        setPayslip((current) => ({ ...current, masterPath: selected }));
      }
    } catch (error) {
      addToast("error", String(error));
    }
  }

  async function browsePayslipLksFiles() {
    if (!bridge) {
      addToast("error", "Desktop bridge not ready.");
      return;
    }
    try {
      const raw = await bridgeCall<string>(bridge, "pickLksFiles");
      if (!raw) {
        return;
      }
      const selected = JSON.parse(raw) as string[];
      setPayslip((current) => {
        const unique = new Map(current.lksPaths.map((value) => [value.toLowerCase(), value]));
        selected.forEach((value) => {
          unique.set(value.toLowerCase(), value);
        });
        return { ...current, lksPaths: [...unique.values()] };
      });
    } catch (error) {
      addToast("error", String(error));
    }
  }

  async function browseDirectory() {
    if (!bridge) {
      addToast("error", "Desktop bridge not ready.");
      return;
    }
    try {
      const selected = await bridgeCall<string>(bridge, "pickDirectory");
      if (!selected) {
        return;
      }
      setPayslip((current) => ({ ...current, outputDir: selected }));
    } catch (error) {
      addToast("error", String(error));
    }
  }

  function startLksRun() {
    if (!bridge) {
      addToast("error", "Desktop bridge not ready.");
      return;
    }
    if (!lks.inputPath.trim()) {
      addToast("error", "Select the input Excel file first.");
      return;
    }
    setLks((current) => ({
      ...current,
      running: true,
      outputPath: "",
      generatedInputPath: "",
      summary: null,
      trasByDate: {},
      statusText: "Starting",
      logLines: ["Starting LKS automation..."],
    }));
    bridge.startLks(JSON.stringify({ inputPath: lks.inputPath.trim() }));
  }

  function startPayslipRun() {
    if (!bridge) {
      addToast("error", "Desktop bridge not ready.");
      return;
    }
    if (!payslip.calcPath.trim()) {
      addToast("error", "Select the calculation workbook.");
      return;
    }
    if (!payslip.masterPath.trim()) {
      addToast("error", "Select the worker master file.");
      return;
    }
    if (!payslip.salaryMonth.trim()) {
      addToast("error", "Enter the salary month.");
      return;
    }
    if (!payslip.paymentDate.trim()) {
      addToast("error", "Select the payment date.");
      return;
    }
    setPayslip((current) => ({
      ...current,
      running: true,
      lastOutputDir: "",
      summary: null,
      logLines: ["Starting payslip generation..."],
    }));

    bridge.startPayslip(
      JSON.stringify({
        calcPath: payslip.calcPath.trim(),
        masterPath: payslip.masterPath.trim(),
        outputDir: payslip.outputDir.trim(),
        salaryMonth: payslip.salaryMonth.trim(),
        paymentDate: payslip.paymentDate.trim(),
        lksPaths: payslip.lksPaths,
      }),
    );
  }

  function respondAppendConfirmation(answer: boolean) {
    setAppendOpen(false);
    bridge?.respondAppendConfirmation(answer);
  }

  const activeMeta = moduleMeta[activeModule];
  const lksSummaryMetrics = useMemo(
    () => buildMetricItems(lksSummarySchema, lks.summary),
    [lks.summary],
  );
  const payslipSummaryMetrics = useMemo(
    () => buildPayslipMetricItems(payslip.summary),
    [payslip.summary],
  );

  return (
    <>
      <div className="flex h-full bg-background">
        <aside className="flex w-[232px] shrink-0 flex-col border-r border-border bg-card px-4 py-4">
          <div className="space-y-1">
            <p className="eyebrow">TNB Workspace</p>
          </div>

          <Separator className="my-4" />

          <nav className="space-y-2">
            {Object.entries(moduleMeta).map(([key, value]) => (
              <button
                key={key}
                type="button"
                onClick={() => setActiveModule(key as ModuleKey)}
                className={cn(
                  "w-full rounded-2xl border px-3.5 py-2.5 text-left transition-colors",
                  activeModule === key
                    ? "border-primary/20 bg-accent text-foreground"
                    : "border-transparent bg-transparent hover:bg-muted",
                )}
              >
                <div className="space-y-0.5">
                  <p className="eyebrow">{value.label}</p>
                  <p className="text-sm font-semibold">{value.title}</p>
                </div>
              </button>
            ))}
          </nav>

          <div className="mt-auto">
            <p className="text-xs text-muted-foreground">Version {version}</p>
          </div>
        </aside>

        <main className="flex min-w-0 flex-1 flex-col overflow-hidden px-4 py-4 xl:px-5">
          <header className="mb-3 flex items-start justify-between gap-4">
            <div className="space-y-2">
              <p className="eyebrow">{activeMeta.label}</p>
              <div className="space-y-1">
                <h2 className="text-2xl font-semibold tracking-tight">{activeMeta.title}</h2>
                <p className="max-w-3xl text-sm text-muted-foreground">{activeMeta.description}</p>
              </div>
            </div>
            <div className="flex items-start gap-3">
              <Button
                variant="secondary"
                size="icon"
                onClick={() => setInfoOpen(true)}
                aria-label="Information"
                title="Information"
              >
                i
              </Button>
            </div>
          </header>

          <div className="min-h-0 flex-1 overflow-hidden">
            {activeModule === "lks" ? (
              <div className="grid h-full grid-cols-1 gap-3 xl:grid-cols-[minmax(420px,0.95fr)_minmax(480px,1.05fr)]">
                <div className="flex min-h-0 flex-col gap-3">
                  <Card>
                    <CardHeader>
                      <CardTitle>Source Workbook</CardTitle>
                      <CardDescription>Select the technician workbook and start the LKS run.</CardDescription>
                    </CardHeader>
                    <CardContent className="space-y-3">
                      <Field label="Input Excel file">
                        <div className="grid grid-cols-[minmax(0,1fr)_auto] gap-3">
                          <Input value={lks.inputPath} readOnly placeholder="Select the technician workbook" />
                          <Button variant="secondary" onClick={() => browseFile("lksInput")} disabled={lks.running}>
                            Browse
                          </Button>
                        </div>
                      </Field>
                      <div className="flex flex-wrap gap-3">
                        <Button onClick={startLksRun} disabled={lks.running}>
                          Run LKS Automation
                        </Button>
                      </div>
                    </CardContent>
                  </Card>

                  <Card className="min-h-0 flex flex-1 flex-col overflow-hidden">
                    <CardHeader>
                      <div className="space-y-1">
                        <CardTitle>Run Log</CardTitle>
                        <CardDescription>Live activity messages from the Python workflow.</CardDescription>
                      </div>
                    </CardHeader>
                    <CardContent className="min-h-0 flex-1">
                      <ScrollArea className="h-full rounded-2xl border border-border bg-muted/40 p-3.5">
                        <pre className="font-mono text-[13px] leading-6 text-slate-700 whitespace-pre-wrap">
                          {lks.logLines.join("\n")}
                        </pre>
                      </ScrollArea>
                    </CardContent>
                  </Card>

                  <Card>
                    <CardHeader>
                      <CardTitle>Output Actions</CardTitle>
                      <CardDescription>Open the generated workbook, the raw processed input, or the result folder.</CardDescription>
                    </CardHeader>
                    <CardContent className="grid grid-cols-1 gap-2.5 md:grid-cols-3">
                      <Button
                        variant="secondary"
                        disabled={lks.running || !lks.outputPath}
                        onClick={() => bridge?.openPath(lks.outputPath)}
                      >
                        Open LKS
                      </Button>
                      <Button
                        variant="secondary"
                        disabled={lks.running || !lks.generatedInputPath}
                        onClick={() => bridge?.openPath(lks.generatedInputPath)}
                      >
                        Open Raw Data
                      </Button>
                      <Button
                        variant="secondary"
                        disabled={lks.running || !lks.outputPath}
                        onClick={() => bridge?.openPath(directoryFromPath(lks.outputPath))}
                      >
                        Open Folder
                      </Button>
                    </CardContent>
                  </Card>
                </div>

                <div className="flex min-h-0 flex-col gap-3">
                  <Card className="min-h-0 flex flex-1 flex-col overflow-hidden">
                    <CardHeader>
                      <CardTitle>Run Summary</CardTitle>
                      <CardDescription>Key totals from the current LKS run.</CardDescription>
                    </CardHeader>
                    <CardContent className="flex flex-1 flex-col gap-3">
                      <MetricGrid items={lksSummaryMetrics} />
                      {Object.keys(lks.trasByDate).length > 0 && (
                        <div className="space-y-2.5 rounded-2xl border border-border bg-muted/40 p-3">
                          <div>
                            <p className="text-sm font-semibold">TRAS by date</p>
                            <p className="text-sm text-muted-foreground">These rows were excluded from the LKS output.</p>
                          </div>
                          <div className="grid gap-1.5 md:grid-cols-2">
                            {Object.entries(lks.trasByDate).map(([label, value]) => (
                              <div
                                key={label}
                                className="flex items-center justify-between rounded-xl bg-background px-2.5 py-1.5"
                              >
                                <span className="text-xs text-muted-foreground">{label}</span>
                                <span className="text-xs font-semibold">{value}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                    </CardContent>
                  </Card>
                </div>
              </div>
            ) : (
              <div className="grid h-full grid-cols-1 gap-3 xl:grid-cols-[minmax(440px,1fr)_minmax(480px,1fr)]">
                <div className="flex min-h-0 flex-col gap-3">
                  <Card className="min-h-0 flex flex-1 flex-col overflow-hidden">
                    <CardHeader>
                      <CardTitle>Generation Inputs</CardTitle>
                      <CardDescription>Use the calculation workbook directly or attach LKS files for auto-fill.</CardDescription>
                    </CardHeader>
                    <CardContent className="space-y-3">
                      <Field label="Calculation workbook">
                        <div className="grid grid-cols-[minmax(0,1fr)_auto] gap-3">
                          <Input value={payslip.calcPath} readOnly placeholder="Select the filled TNB calculation workbook" />
                          <Button variant="secondary" onClick={() => browseFile("calc")} disabled={payslip.running}>
                            Browse
                          </Button>
                        </div>
                      </Field>

                      <Field label="Worker master file">
                        <div className="grid grid-cols-[minmax(0,1fr)_auto] gap-3">
                          <Input value={payslip.masterPath} readOnly placeholder="Select the worker master workbook" />
                          <Button variant="secondary" onClick={() => browseFile("master")} disabled={payslip.running}>
                            Browse
                          </Button>
                        </div>
                      </Field>

                      <Field label="LKS files (optional)">
                        <div className="grid grid-cols-[minmax(0,1fr)_auto] gap-3">
                          <Input
                            value={
                              payslip.lksPaths.length === 0
                                ? ""
                                : payslip.lksPaths.length === 1
                                  ? payslip.lksPaths[0]
                                  : `${payslip.lksPaths.length} files selected`
                            }
                            readOnly
                            placeholder="No LKS files selected"
                          />
                          <Button variant="secondary" onClick={browsePayslipLksFiles} disabled={payslip.running}>
                            Add Files
                          </Button>
                        </div>
                        {payslip.lksPaths.length > 0 && (
                          <div className="mt-2 flex flex-wrap gap-2">
                            {payslip.lksPaths.map((value) => (
                              <Badge key={value} variant="secondary" className="max-w-full truncate px-3 py-1.5 text-xs">
                                {basename(value)}
                              </Badge>
                            ))}
                          </div>
                        )}
                        <div className="mt-2">
                          <Button
                            variant="ghost"
                            onClick={() => setPayslip((current) => ({ ...current, lksPaths: [] }))}
                            disabled={payslip.running || payslip.lksPaths.length === 0}
                          >
                            Clear LKS Files
                          </Button>
                        </div>
                      </Field>

                      <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                        <Field label="Salary month">
                          <Input
                            value={payslip.salaryMonth}
                            onChange={(event) =>
                              setPayslip((current) => ({ ...current, salaryMonth: event.target.value }))
                            }
                            placeholder="Example: May 2026"
                            disabled={payslip.running}
                          />
                        </Field>
                        <Field label="Payment date">
                          <Input
                            type="date"
                            className="calendar-input"
                            value={payslip.paymentDate}
                            onChange={(event) =>
                              setPayslip((current) => ({ ...current, paymentDate: event.target.value }))
                            }
                            disabled={payslip.running}
                          />
                        </Field>
                      </div>

                      <Field label="Output folder">
                        <div className="grid grid-cols-[minmax(0,1fr)_auto] gap-3">
                          <Input value={payslip.outputDir} readOnly placeholder="Select the output folder" />
                          <Button variant="secondary" onClick={browseDirectory} disabled={payslip.running}>
                            Browse
                          </Button>
                        </div>
                      </Field>

                      <div className="grid grid-cols-1 gap-3 pt-2 md:grid-cols-2">
                        <Button onClick={startPayslipRun} disabled={payslip.running}>
                          Generate Payslips
                        </Button>
                        <Button
                          variant="secondary"
                          disabled={payslip.running || !payslip.lastOutputDir}
                          onClick={() => bridge?.openPath(payslip.lastOutputDir)}
                        >
                          Open Output Folder
                        </Button>
                      </div>
                    </CardContent>
                  </Card>
                </div>

                <div className="flex min-h-0 flex-col gap-3">
                  <Card className="min-h-0 flex flex-1 flex-col overflow-hidden">
                    <CardHeader>
                      <CardTitle>Generation Summary</CardTitle>
                      <CardDescription>Results from the current payslip run.</CardDescription>
                    </CardHeader>
                    <CardContent className="min-h-0 flex-1 space-y-4 overflow-auto">
                      <MetricGrid items={payslipSummaryMetrics} />

                      {payslip.summary && (
                        <div className="space-y-3 rounded-2xl border border-border bg-muted/40 p-3.5">
                          <div>
                            <p className="text-sm font-semibold">Generation details</p>
                            <p className="text-sm text-muted-foreground">
                              Output folder, warnings, PDF failures, and generated calculation workbook.
                            </p>
                          </div>

                          <div className="grid gap-2">
                            <DetailRow label="Output folder" value={payslip.summary.outputDir} />
                            {payslip.summary.calculationWorkbookPath ? (
                              <DetailRow
                                label="Calculation workbook"
                                value={payslip.summary.calculationWorkbookPath}
                              />
                            ) : null}
                            {payslip.summary.claimSummary ? (
                              <DetailRow
                                label="LKS files used"
                                value={String(payslip.summary.claimSummary.sourceFiles)}
                              />
                            ) : null}
                            {payslip.summary.warnings.map((warning) => (
                              <DetailRow key={`warning-${warning}`} label="Warning" value={warning} />
                            ))}
                            {payslip.summary.pdfFailures.map((failure) => (
                              <DetailRow key={`pdf-${failure}`} label="PDF failure" value={failure} />
                            ))}
                          </div>
                        </div>
                      )}
                    </CardContent>
                  </Card>

                  <Card className="min-h-0 flex flex-1 flex-col overflow-hidden">
                    <CardHeader className="flex-row items-center justify-between space-y-0">
                      <div className="space-y-1">
                        <CardTitle>Generation Log</CardTitle>
                        <CardDescription>Live activity messages from the payslip workflow.</CardDescription>
                      </div>
                      <Button
                        variant="ghost"
                        onClick={() =>
                          setPayslip((current) => ({
                            ...current,
                            logLines: ["No generation started yet."],
                          }))
                        }
                      >
                        Clear
                      </Button>
                    </CardHeader>
                    <CardContent className="min-h-0 flex-1">
                      <ScrollArea className="h-full rounded-2xl border border-border bg-muted/40 p-3.5">
                        <pre className="font-mono text-[13px] leading-6 text-slate-700 whitespace-pre-wrap">
                          {payslip.logLines.join("\n")}
                        </pre>
                      </ScrollArea>
                    </CardContent>
                  </Card>
                </div>
              </div>
            )}
          </div>
        </main>
      </div>

      <Dialog open={infoOpen} onOpenChange={setInfoOpen}>
        <DialogContent>
          <DialogHeader>
            <p className="eyebrow">App Information</p>
            <DialogTitle>Support and version details</DialogTitle>
            <DialogDescription>
              Keep this dialog simple. Version and support details belong here, not in the main workspace header.
            </DialogDescription>
          </DialogHeader>
          <div className="grid gap-3">
            <InfoRow label="Version" value={version} />
            <InfoRow label="Support Email" value={supportEmail} />
            <InfoRow label="Support Phone" value={supportPhone} />
          </div>
          <DialogFooter className="gap-2 pt-2">
            <Button variant="secondary" onClick={() => bridge?.checkUpdates(activeModule)}>
              Check Updates
            </Button>
            <Button onClick={() => setInfoOpen(false)}>Done</Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <Dialog open={appendOpen} onOpenChange={setAppendOpen}>
        <DialogContent>
          <DialogHeader>
            <p className="eyebrow">Existing Data Found</p>
            <DialogTitle>Append to the current template?</DialogTitle>
            <DialogDescription>
              The template already has {appendPrompt.existing} SOs. The current run will add {appendPrompt.incoming} new SOs.
            </DialogDescription>
          </DialogHeader>
          <DialogFooter className="gap-2 pt-2">
            <Button variant="secondary" onClick={() => respondAppendConfirmation(false)}>
              Cancel Run
            </Button>
            <Button onClick={() => respondAppendConfirmation(true)}>Continue</Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <div className="fixed bottom-6 right-6 z-50 flex flex-col gap-3">
        {toasts.map((toast) => (
          <div
            key={toast.id}
            className={cn(
              "w-[320px] rounded-2xl border bg-card px-4 py-3 shadow-panel",
              toast.level === "error" && "border-destructive/20",
              toast.level === "success" && "border-success/20",
            )}
          >
            <p className="text-sm font-semibold">
              {toast.level === "error" ? "Action failed" : toast.level === "success" ? "Done" : "Notice"}
            </p>
            <p className="mt-1 text-sm text-muted-foreground">{toast.message}</p>
          </div>
        ))}
      </div>
    </>
  );
}

function Field({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="space-y-2">
      <label className="text-sm font-medium text-foreground">{label}</label>
      {children}
    </div>
  );
}

function MetricGrid({
  items,
}: {
  items: Array<{ key: string; label: string; value: string }>;
}) {
  return (
    <div className="grid grid-cols-2 gap-2 lg:grid-cols-3">
      {items.map((item) => (
        <div key={item.key} className="min-w-0 rounded-2xl border border-border bg-background px-3 py-2.5">
          <p className="text-[10px] font-semibold uppercase tracking-[0.14em] text-muted-foreground">
            {item.label}
          </p>
          <p className="mt-1 text-base font-semibold tracking-tight text-foreground break-words">
            {item.value}
          </p>
        </div>
      ))}
    </div>
  );
}

function buildMetricItems(
  schema: Array<{ key: string; label: string }>,
  summary: Record<string, string | number> | null,
) {
  return schema.map((item) => ({
    key: item.key,
    label: item.label,
    value:
      summary && Object.prototype.hasOwnProperty.call(summary, item.key)
        ? String(summary[item.key])
        : "--",
  }));
}

function buildPayslipMetricItems(summary: PayslipRunResult | null) {
  const items: Array<{ key: string; label: string; value: string }> = [
    {
      key: "total-sos",
      label: "Total SOs",
      value: summary?.claimSummary ? String(summary.claimSummary.countedRows) : "--",
    },
  ];

  const fileSummaries = summary?.claimSummary?.fileSummaries ?? [];
  fileSummaries.forEach((fileSummary, index) => {
    items.push({
      key: `${fileSummary.fileName}-${index}`,
      label: simplifyLksFileLabel(fileSummary.fileName, index),
      value: String(fileSummary.countedRows),
    });
  });

  return items;
}

function simplifyLksFileLabel(fileName: string, index: number) {
  const match = fileName.match(/lks\s*(\d+)/i);
  if (match) {
    return `LKS ${match[1]} SOs`;
  }
  return `LKS ${index + 1} SOs`;
}

function DetailRow({ label, value }: { label: string; value: string }) {
  return (
    <div className="rounded-xl bg-background px-3 py-3">
      <p className="text-xs font-semibold uppercase tracking-[0.14em] text-muted-foreground">{label}</p>
      <p className="mt-1 break-all text-sm font-medium text-foreground">{value}</p>
    </div>
  );
}

function InfoRow({ label, value }: { label: string; value: string }) {
  return (
    <div className="rounded-[20px] border border-border bg-muted/40 px-4 py-3.5">
      <p className="text-xs font-semibold uppercase tracking-[0.14em] text-muted-foreground">{label}</p>
      <p className="mt-1 text-sm font-medium text-foreground">{value}</p>
    </div>
  );
}

function bridgeCall<T>(
  bridge: DesktopBridge,
  methodName: keyof DesktopBridge,
  ...args: string[]
): Promise<T> {
  return new Promise((resolve, reject) => {
    const candidate = bridge[methodName];
    if (typeof candidate !== "function") {
      reject(new Error(`Bridge method not found: ${String(methodName)}`));
      return;
    }

    try {
      (candidate as (...callArgs: unknown[]) => void)(...args, (value: T) => resolve(value));
    } catch (error) {
      reject(error);
    }
  });
}

function trimPlaceholder(values: string[], placeholder: string) {
  return values.length === 1 && values[0] === placeholder ? [] : values;
}

function appendUniqueLog(values: string[], placeholder: string, nextLine: string) {
  const base = trimPlaceholder(values, placeholder);
  if (base[base.length - 1] === nextLine) {
    return base;
  }
  return [...base, nextLine];
}

function basename(path: string) {
  const parts = path.split(/[\\/]/);
  return parts[parts.length - 1] || path;
}

function directoryFromPath(path: string) {
  const parts = path.split(/[\\/]/);
  parts.pop();
  return parts.join("\\");
}

function formatMonthLabel(value: Date) {
  return value.toLocaleDateString("en-GB", { month: "long", year: "numeric" });
}

function formatDateInput(value: Date) {
  const year = value.getFullYear();
  const month = String(value.getMonth() + 1).padStart(2, "0");
  const day = String(value.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

export default App;
