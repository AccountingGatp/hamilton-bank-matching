# Chase and ResMan Matching Statement

This app accepts a Chase statement PDF, a ResMan bank register XLSX, a prior
finalized ResMan reconciliation PDF, and an AP Aging Detail XLSX. It produces an
Excel matching statement with the four SOP sections, match detail, review
candidates, and source controls.

## Deploy through GitHub

1. Extract this ZIP and commit its contents, including `.gitignore` and
   `.python-version`, at the root of a GitHub repository.
2. In Vercel, choose **Add New → Project** and import that repository. Leave the
   Root Directory at the repository root and leave the build settings at their
   defaults. The Framework Preset is **Other** and the project configuration
   routes the upload page and reconciliation requests to a Python Function.
3. Deploy and open the deployment URL. No environment variables or local
   Poppler installation are needed. Subsequent pushes to the linked production
   branch trigger deployments.
4. For team use with real statements, restrict access to the deployment URL
   using Vercel Authentication with Deployment Protection set to All Deployments. A private GitHub repository alone does
   not restrict access to the deployed URL. Check protection on the production
   domain where the team will upload files.

Commit application files only. Upload statements and exports through the form.
Files are processed in an invocation-specific temporary directory, removed after
processing, and returned directly to the browser without an application database.

## Input limits

Use the supported text-based Chase and ResMan export layouts, one account per
run. The Chase statement determines the current dates. The ResMan register can
cover a wider range and can show a reporting month or an explicit date range.
It must cover the interval from the day after the prior finalized reconciliation
through the Chase ending date. The workbook uses that Chase ending date as the
books cutoff. The AP Aging report supplies invoice suggestions when its snapshot
is dated on or after the statement cutoff and belongs to the same property.

The upload form accepts up to 4 MB combined to allow
room for multipart headers under Vercel's default Function payload limit.
The PDF extractor preserves words and punctuation, and its column spacing may
differ from the local Poppler version. Review provisional PayLease groups and
unresolved same-amount candidates using the workbook's Review Candidates tab.
