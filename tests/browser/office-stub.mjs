// Replaces office.js in every frame. Holds fixture data only; any Office
// member the add-in reads that is not modelled here throws.

export function officeStub(fixture) {
  const top = window.top;
  const strict = (name, target) =>
    new Proxy(target, {
      get(obj, prop) {
        if (typeof prop === "symbol" || prop in obj || prop === "then" || prop === "toJSON") {
          return obj[prop];
        }
        throw new Error(`office stub: ${name}.${String(prop)} is not modelled`);
      },
    });
  const done = (value) => ({ status: "succeeded", value });

  // A dialog frame only gets what the dialog API documents as callable there,
  // plus the display language the dialogs read.
  const isDialog = window !== top;
  let context;
  if (isDialog) {
    context = strict("context", {
      displayLanguage: fixture.language,
      ui: strict("ui", {
        messageParent: (message) => top.__officeState.deliver(message),
        addHandlerAsync: (type, handler, cb) => {
          if (type !== "dialogParentMessageReceived") {
            throw new Error(`office stub: dialog handler for ${type} is not modelled`);
          }
          top.__officeState.dialog.childHandler = handler;
          cb(done());
        },
      }),
      requirements: strict("requirements", { isSetSupported: () => true }),
    });
  } else {
    const state = {
      roaming: { ...fixture.roaming },
      saved: null,
      dialogUrls: [],
      childMessages: [],
      recipients: { to: [], cc: [], bcc: [] },
      categories: [...fixture.categories],
      masterCategories: [...fixture.masterCategories],
      dialog: null,
      deliver(message) {
        if (!state.dialog) {
          throw new Error("office stub: messageParent without an open dialog");
        }
        state.dialog.handler({ message });
      },
    };
    window.__officeState = state;

    const recipients = (kind) =>
      strict(kind, {
        getAsync: (cb) => cb(done([...state.recipients[kind]])),
        addAsync: (list, cb) => {
          state.recipients[kind].push(...list);
          cb(done());
        },
      });
    const message = fixture.item;
    const mailboxAtLeast = (version) => {
      const [major, minor] = (fixture.requirementSets.Mailbox || "1.0").split(".").map(Number);
      const [wantMajor, wantMinor] = version.split(".").map(Number);
      return major > wantMajor || (major === wantMajor && minor >= wantMinor);
    };
    const item =
      message.mode === "compose"
        ? strict("item", { to: recipients("to"), cc: recipients("cc"), bcc: recipients("bcc") })
        : strict("item", {
            from: message.from,
            to: message.to,
            cc: message.cc,
            bcc: message.bcc,
            subject: message.subject,
            dateTimeCreated: new Date(message.dateTimeCreated),
            // holdBody keeps the body read pending until the test calls releaseBody().
            body: strict("body", {
              getAsync: (type, cb) => {
                if (!message.holdBody) {
                  cb(done(message.body));
                  return;
                }
                state.bodyRead = true;
                state.releaseBody = () => cb(done(message.body));
              },
            }),
            // Outlook offers item.categories from Mailbox 1.8 on.
            ...(mailboxAtLeast("1.8") && {
              categories: strict("categories", {
                getAsync: (cb) =>
                  cb(
                    message.categoriesFail
                      ? { status: "failed", error: { code: 9002, message: "internal error" } }
                      : done(state.categories.map((displayName) => ({ displayName }))),
                  ),
                // Outlook refuses a category missing from the master list with 9044.
                addAsync: (names, cb) => {
                  if (!names.every((n) => state.masterCategories.includes(n))) {
                    cb({ status: "failed", error: { code: 9044, message: "not in master list" } });
                    return;
                  }
                  state.categories.push(...names);
                  cb(done());
                },
              }),
            }),
          });

    // Only the requirement sets the fixture names are modelled.
    const isSetSupported = (name, version) => {
      if (!(name in fixture.requirementSets)) {
        throw new Error(`office stub: requirement set ${name} is not modelled`);
      }
      const [major, minor] = fixture.requirementSets[name].split(".").map(Number);
      const [wantMajor, wantMinor] = version.split(".").map(Number);
      return major > wantMajor || (major === wantMajor && minor >= wantMinor);
    };

    context = strict("context", {
      displayLanguage: fixture.language,
      requirements: strict("requirements", { isSetSupported }),
      roamingSettings: strict("roamingSettings", {
        get: (key) => state.roaming[key],
        set: (key, value) => {
          state.roaming[key] = value;
        },
        saveAsync: (cb) => {
          state.saved = { ...state.roaming };
          cb(done());
        },
      }),
      mailbox: strict("mailbox", {
        item,
        masterCategories: strict("masterCategories", {
          addAsync: (list, cb) => {
            state.masterCategories.push(...list.map((c) => c.displayName));
            cb(done());
          },
        }),
      }),
      ui: strict("ui", {
        displayDialogAsync: (url, options, cb) => {
          if (state.dialog) {
            throw new Error("office stub: a second dialog while one is open");
          }
          state.dialogUrls.push(url);
          const frame = document.createElement("iframe");
          frame.id = "office-dialog";
          frame.src = url;
          frame.style.cssText =
            "position:fixed;inset:0;width:100%;height:100%;background:#fff;border:0;z-index:99999";
          document.body.appendChild(frame);
          state.dialog = { frame, handler: null, childHandler: null };
          cb(
            done(
              strict("dialog", {
                addEventHandler: (type, handler) => {
                  state.dialog.handler = handler;
                },
                // Real Office drops a message the dialog is not yet listening for.
                messageChild: (text) => {
                  if (!state.dialog.childHandler) {
                    throw new Error("office stub: messageChild before the dialog listens");
                  }
                  state.childMessages.push(text.length);
                  state.dialog.childHandler({ message: text, origin: location.origin });
                },
                close: () => {
                  frame.remove();
                  state.dialog = null;
                },
              }),
            ),
          );
        },
      }),
    });
  }

  const office = {
    context,
    EventType: strict("EventType", { DialogParentMessageReceived: "dialogParentMessageReceived" }),
    AsyncResultStatus: strict("AsyncResultStatus", { Succeeded: "succeeded", Failed: "failed" }),
    MailboxEnums: strict("MailboxEnums", {
      CategoryColor: strict("CategoryColor", { Preset0: "preset0" }),
    }),
    initialize: null,
  };
  window.Office = strict("Office", office);
  window.Microsoft = {
    Office: { WebExtension: { EventType: { DialogMessageReceived: "dialogMessageReceived" } } },
  };
  // office.js calls Office.initialize once the host handshake is done.
  window.addEventListener("DOMContentLoaded", () => {
    if (typeof office.initialize !== "function") {
      throw new Error("office stub: the page set no Office.initialize");
    }
    office.initialize("documentOpened");
  });
}
