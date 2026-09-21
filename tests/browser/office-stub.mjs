// Replaces office.js in every frame. Holds fixture data only; any Office
// member the add-in reads that is not modelled here throws.

export function officeStub(fixture) {
  const top = window.top;
  const strict = (name, target) =>
    new Proxy(target, {
      get(obj, prop) {
        if (typeof prop === 'symbol' || prop in obj || prop === 'then' || prop === 'toJSON') {
          return obj[prop];
        }
        throw new Error(`office stub: ${name}.${String(prop)} is not modelled`);
      },
    });
  const done = (value) => ({ status: 'succeeded', value });

  // A dialog frame only gets what the dialog API documents as callable there,
  // plus the display language the dialogs read.
  const isDialog = window !== top;
  let context;
  if (isDialog) {
    context = strict('context', {
      displayLanguage: fixture.language,
      ui: strict('ui', {
        messageParent: (message) => top.__officeState.deliver(message),
      }),
      requirements: strict('requirements', { isSetSupported: () => true }),
    });
  } else {
    const state = {
      roaming: { ...fixture.roaming },
      saved: null,
      dialogUrls: [],
      recipients: { to: [], cc: [], bcc: [] },
      categories: [...fixture.categories],
      masterCategories: [...fixture.masterCategories],
      dialog: null,
      deliver(message) {
        if (!state.dialog) {
          throw new Error('office stub: messageParent without an open dialog');
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
    const item =
      message.mode === 'compose'
        ? strict('item', { to: recipients('to'), cc: recipients('cc'), bcc: recipients('bcc') })
        : strict('item', {
            from: message.from,
            to: message.to,
            cc: message.cc,
            bcc: message.bcc,
            subject: message.subject,
            dateTimeCreated: new Date(message.dateTimeCreated),
            body: strict('body', { getAsync: (type, cb) => cb(done(message.body)) }),
            categories: strict('categories', {
              getAsync: (cb) => cb(done(state.categories.map((displayName) => ({ displayName })))),
              // Outlook refuses a category missing from the master list with 9044.
              addAsync: (names, cb) => {
                if (!names.every((n) => state.masterCategories.includes(n))) {
                  cb({ status: 'failed', error: { code: 9044, message: 'not in master list' } });
                  return;
                }
                state.categories.push(...names);
                cb(done());
              },
            }),
          });

    context = strict('context', {
      displayLanguage: fixture.language,
      roamingSettings: strict('roamingSettings', {
        get: (key) => state.roaming[key],
        set: (key, value) => {
          state.roaming[key] = value;
        },
        saveAsync: (cb) => {
          state.saved = { ...state.roaming };
          cb(done());
        },
      }),
      mailbox: strict('mailbox', {
        item,
        restUrl: fixture.restUrl,
        getCallbackTokenAsync: (options, cb) => cb(done(fixture.callbackToken)),
        masterCategories: strict('masterCategories', {
          addAsync: (list, cb) => {
            state.masterCategories.push(...list.map((c) => c.displayName));
            cb(done());
          },
        }),
      }),
      ui: strict('ui', {
        displayDialogAsync: (url, options, cb) => {
          if (state.dialog) {
            throw new Error('office stub: a second dialog while one is open');
          }
          state.dialogUrls.push(url);
          const frame = document.createElement('iframe');
          frame.id = 'office-dialog';
          frame.src = url;
          frame.style.cssText = 'position:fixed;inset:0;width:100%;height:100%;background:#fff;border:0;z-index:99999';
          document.body.appendChild(frame);
          state.dialog = { frame, handler: null };
          cb(
            done(
              strict('dialog', {
                addEventHandler: (type, handler) => {
                  state.dialog.handler = handler;
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
    AsyncResultStatus: strict('AsyncResultStatus', { Succeeded: 'succeeded', Failed: 'failed' }),
    MailboxEnums: strict('MailboxEnums', { CategoryColor: strict('CategoryColor', { Preset0: 'preset0' }) }),
    initialize: null,
  };
  window.Office = strict('Office', office);
  window.Microsoft = {
    Office: { WebExtension: { EventType: { DialogMessageReceived: 'dialogMessageReceived' } } },
  };
  // office.js calls Office.initialize once the host handshake is done.
  window.addEventListener('DOMContentLoaded', () => {
    if (typeof office.initialize !== 'function') {
      throw new Error('office stub: the page set no Office.initialize');
    }
    office.initialize('documentOpened');
  });
}
