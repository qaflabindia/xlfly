// Task pane: trust, module editor, catalog of functions and commands, log.

import type { App } from "../app.ts";
import type { Module } from "../shared/codestore-format.ts";

const STARTER = `import xlfly


@xlfly.func
def double(x):
    """=XLFLY.CALL("DOUBLE", A1)"""
    return x * 2


@xlfly.command
async def hello():
    """Writes a greeting to A1 of the active sheet."""
    await xlfly.Range("A1").write("Hello from Python")
`;

export function bindTaskpane(app: App, root: HTMLElement): void {
  const $ = <T extends HTMLElement>(id: string) => root.querySelector<T>(`#${id}`)!;
  const status = $("kernel-status");
  const trustBox = $("trust");
  const trustText = $("trust-text");
  const trustButton = $<HTMLButtonElement>("trust-button");
  const moduleSelect = $<HTMLSelectElement>("module-select");
  const editor = $<HTMLTextAreaElement>("editor");
  const saveButton = $<HTMLButtonElement>("save-button");
  const newButton = $<HTMLButtonElement>("new-module");
  const deleteButton = $<HTMLButtonElement>("delete-module");
  const newForm = $<HTMLFormElement>("new-module-form");
  const newName = $<HTMLInputElement>("new-module-name");
  const functions = $("functions");
  const commands = $("commands");
  const log = $("log");
  $("version").textContent = `v${app.version}`;

  let drafts: Module[] = [];
  let selected = 0;
  let dirty = false;

  const loadDrafts = () => {
    drafts = (app.doc?.modules ?? []).map((m) => ({ ...m }));
    if (drafts.length === 0) drafts.push({ name: "main", source: STARTER });
    selected = Math.min(selected, drafts.length - 1);
    dirty = false;
  };

  const renderModules = () => {
    moduleSelect.replaceChildren(
      ...drafts.map((m, i) => {
        const option = document.createElement("option");
        option.value = String(i);
        option.textContent = m.name;
        option.selected = i === selected;
        return option;
      }),
    );
    editor.value = drafts[selected]?.source ?? "";
    editor.readOnly = !app.trusted;
  };

  const render = () => {
    status.textContent = app.kernelStatus;
    status.dataset.state = app.kernelStatus;
    status.title = app.kernelDetail;

    trustBox.dataset.trusted = String(app.trusted);
    trustText.textContent = app.trusted
      ? "Python code in this workbook is trusted on this computer."
      : "Python code in this workbook is not trusted and will not run. Review the code below, then trust it.";
    trustButton.textContent = app.trusted ? "Revoke trust" : "Trust this workbook";
    saveButton.disabled = !app.trusted;
    newButton.disabled = !app.trusted;
    deleteButton.disabled = !app.trusted || drafts.length <= 1;
    editor.readOnly = !app.trusted;

    const catalog = app.kernel.catalog;
    functions.replaceChildren(
      ...Object.entries(catalog.functions).map(([name, info]) => {
        const li = document.createElement("li");
        li.innerHTML = `<code></code><span></span>`;
        li.querySelector("code")!.textContent = `=XLFLY.CALL("${name}"${info.signature === "()" ? "" : ", …"})`;
        li.querySelector("span")!.textContent = info.doc;
        return li;
      }),
    );
    commands.replaceChildren(
      ...Object.entries(catalog.commands).map(([name, info]) => {
        const li = document.createElement("li");
        const button = document.createElement("button");
        button.textContent = `Run ${name}`;
        button.onclick = () => void app.runCommand(name);
        const doc = document.createElement("span");
        doc.textContent = info.doc;
        li.append(button, doc);
        return li;
      }),
    );
    log.textContent = app.logLines.join("\n");
    log.scrollTop = log.scrollHeight;
  };

  moduleSelect.onchange = () => {
    drafts[selected].source = editor.value;
    selected = Number(moduleSelect.value);
    renderModules();
  };
  editor.oninput = () => {
    dirty = true;
  };
  editor.onkeydown = (event) => {
    if (event.key === "Tab") {
      event.preventDefault();
      editor.setRangeText("    ", editor.selectionStart, editor.selectionEnd, "end");
      dirty = true;
    }
  };
  // window.prompt() is unavailable in Office add-in web views, so the name is asked for inline.
  newButton.onclick = () => {
    newForm.hidden = false;
    newName.value = `module${drafts.length + 1}`;
    newName.focus();
    newName.select();
  };
  $("new-module-cancel").onclick = () => {
    newForm.hidden = true;
  };
  newName.onkeydown = (event) => {
    if (event.key === "Escape") newForm.hidden = true;
  };
  newForm.onsubmit = (event) => {
    event.preventDefault();
    const name = newName.value.trim();
    if (!/^[A-Za-z_]\w*$/.test(name) || drafts.some((m) => m.name === name)) {
      app.log(`invalid or duplicate module name: ${name}`);
      newName.focus();
      return;
    }
    newForm.hidden = true;
    drafts[selected].source = editor.value;
    drafts.push({ name, source: "import xlfly\n" });
    selected = drafts.length - 1;
    dirty = true;
    renderModules();
    render();
  };
  deleteButton.onclick = () => {
    if (drafts.length <= 1) return;
    drafts.splice(selected, 1);
    selected = 0;
    dirty = true;
    renderModules();
    render();
  };
  saveButton.onclick = async () => {
    drafts[selected].source = editor.value;
    saveButton.disabled = true;
    try {
      await app.saveModules(drafts.map((m) => ({ ...m })));
      dirty = false;
    } catch (error) {
      app.log(`save failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      saveButton.disabled = !app.trusted;
    }
  };
  trustButton.onclick = async () => {
    trustButton.disabled = true;
    try {
      if (app.trusted) await app.revokeTrust();
      else await app.trustWorkbook();
    } catch (error) {
      app.log(`trust change failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      trustButton.disabled = false;
    }
  };
  $("restart-button").onclick = () => {
    app.kernel.restart("restarted by user");
    void app.reload();
  };

  let lastDoc = app.doc;
  app.subscribe(() => {
    if (app.doc !== lastDoc && !dirty) {
      lastDoc = app.doc;
      loadDrafts();
      renderModules();
    }
    render();
  });

  loadDrafts();
  renderModules();
  render();
}
