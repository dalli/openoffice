// Isolated native-view prototype: synthetic documents and local fake responses only.
const { app, BrowserWindow, WebContentsView } = require('electron')
const { workspaceLayout, paneRects, clampSplitRatio, dividerRect } = require(
  process.env.GENOFFICE_WORKSPACE_PROTOTYPE_MODEL,
)

app
  .whenReady()
  .then(async () => {
    const win = new BrowserWindow({
      width: 1280,
      height: 720,
      show: false,
      webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false },
    })
    await win.loadFile(require('node:path').join(__dirname, 'workspace-layout-prototype.html'))
    await win.webContents.executeJavaScript(
      require('node:fs').readFileSync(process.env.GENOFFICE_WORKSPACE_PROTOTYPE_DIVIDER, 'utf8'),
    )
    const views = []
    for (const label of ['Synthetic document A', 'Synthetic document B']) {
      const view = new WebContentsView({
        webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false },
      })
      win.contentView.addChildView(view)
      await view.webContents.loadURL(
        `data:text/html,${encodeURIComponent(`<h1>${label}</h1><p>Unsaved synthetic material.</p>`)}`,
      )
      views.push(view)
    }
    globalThis.workspacePrototype = {
      async arrange(input) {
        win.setContentSize(input.width, input.height)
        const { width, height } = win.getContentBounds()
        const viewport = { width, height, uiScale: input.uiScale }
        const layout = workspaceLayout(viewport, input.panel, input.orientation)
        const ratio = input.orientation
          ? clampSplitRatio(input.ratio, layout.documents, input.orientation)
          : input.ratio
        const rects = input.orientation
          ? paneRects(layout.documents, input.orientation, ratio)
          : { a: layout.documents, b: layout.documents }
        views[0].setBounds(rects.a)
        views[1].setBounds(rects.b)
        views[1].setVisible(input.orientation !== null)
        const divider = input.orientation
          ? dividerRect({ ...viewport, documents: layout.documents }, input.orientation, ratio)
          : null
        await win.webContents.executeJavaScript(
          `window.arrange(${JSON.stringify({ layout, divider, viewport, orientation: input.orientation, ratio: input.ratio, uiScale: input.uiScale })})`,
        )
        return {
          layout,
          views: views.map((view) => ({ bounds: view.getBounds(), visible: view.getVisible() })),
        }
      },
    }
    win.show()
  })
  .catch((error) => {
    console.error(error)
    app.exit(1)
  })

app.on('window-all-closed', () => app.quit())
