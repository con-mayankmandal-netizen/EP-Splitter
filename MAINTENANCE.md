# EP-Splitter Maintenance Mode

The public GitHub Pages entry point is currently a maintenance page:

- `frontend/index.html`

The working splitter UI has been preserved here:

- `frontend/index.tool.html`

## Restore the tool

To bring the tool back exactly as it was before maintenance mode, copy the preserved tool file over the public entry point:

```bash
cp frontend/index.tool.html frontend/index.html
git add frontend/index.html
git commit -m "Restore EP-Splitter tool"
git push origin main
```

GitHub Pages will deploy from the `frontend` folder through the existing workflow.

## Re-enable maintenance mode

To put the tool back under maintenance later, restore the maintenance-page version of `frontend/index.html` from git history or ask Codex to reapply maintenance mode.
