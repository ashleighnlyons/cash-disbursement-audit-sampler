# Deploy to GitHub Pages

## First deployment

1. Create a new GitHub repository.
2. Add every file and folder from this package to the repository root.
3. Commit the files to the `main` branch.
4. Open **Settings > Pages** in GitHub.
5. Under **Build and deployment**, choose **GitHub Actions**.
6. Open **Actions** and select **Deploy static site to GitHub Pages**.
7. Run the workflow if it did not start automatically.
8. Open the URL shown after the deployment finishes.

Every later push to `main` automatically rebuilds and republishes the website.

## Repository path

The workflow automatically configures a project URL such as:

`https://organization.github.io/cash-disbursement-audit-sampler/`

If the repository is named `organization.github.io`, it builds for the root URL instead.

## Custom domain

Add a custom domain through **Settings > Pages** after the first successful deployment. If the custom domain serves this project at the domain root, remove the **Configure repository base path** step from `.github/workflows/deploy.yml` and rebuild with an empty `PAGES_BASE_PATH`.

## Updating the application

1. Make the change on a branch.
2. Run `npm ci`, `npm run typecheck`, and `npm run build`.
3. Test a known Yardi report and a known AppFolio report.
4. Merge the change into `main`.
5. GitHub Pages publishes the updated static site automatically.

## Privacy model

GitHub Pages sends only the application files—HTML, JavaScript, and CSS—to the browser. After the application loads:

- The selected Excel report is read with the browser File API.
- Workbook parsing and sampling occur in browser memory.
- The output workbook is created in browser memory.
- The browser downloads the output directly.
- No report or generated workbook is sent to GitHub, an API, ChatGPT, Azure, or another server.

The website may be publicly reachable depending on the repository and organization settings. Public reachability does not upload the selected workbook, but it allows anyone with the URL to use the tool and view its client-side code.
