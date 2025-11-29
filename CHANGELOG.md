# Changelog

## 2025-11-29

### Add document generator API endpoint

- Added `/doc-gen` endpoint for further testing of the MongoDB connection and PDF generation;
- It outputs the JSON data into a PDF file at the endpoint;
- PR: [#3 Add doc gen api](https://github.com/kynexi/farmxpert-back-deploy/pull/3)

### Add endpoint for extracting profile data

- Added an API endpoint for extracting profile data based on `OwnerId`;
- This endpoint is currently used to test the MongoDB connection;
- Fixed a mismatch in the naming of `OwnerId` between the codebase and MongoDB.
- PR: [#2 Add extract profile data](https://github.com/kynexi/farmxpert-back-deploy/pull/2)

### Add MongoDB connection

- Refactored codebase from SQL connection to MongoDB;
- Removed unnecessary and outdated code from previous iterations;
- PR: [#1 Add MongoDB connection](https://github.com/kynexi/farmxpert-back-deploy/pull/1)

## 2025-11-28

### Deploy Backend

- Added Vercel deployment for the Flask backend, enabling collaboration with the .NET backend;
- Migrated to a new GitHub repository: https://github.com/kynexi/farmxpert-back-deploy;
- Backend is accessible at https://farmxpert-back-deploy.vercel.app/.

## 2025-11-02

### API reference

- Adds a concise markdown listing of the most important HTTP endpoints, with an example request body and handler locations.
  This results in clearer developer docs.
- Commit: https://github.com/Klavrin/FarmXpert-backend/commit/c9ce2bc883e8ff95cb06d5594d3139a94f4361d5
- Authors: @kynexi

## 2025-10-31

### Improve AI subsidy completion flow

- Improves how subsidy completion is determined by the AI logic to prevent premature or incorrect completion states. This results in more reliable completion checks for subsidy-related operations.
- PR: [#3 Fix ai subsidy completion](https://github.com/Klavrin/FarmXpert-backend/pull/3)
- Authors: @kynexi

### Fix eligibility determination logic

- Corrects the logic used to evaluate user or entity eligibility, addressing incorrect outcomes under certain conditions. This results in more accurate eligibility results; reduces false positives/negatives.
- PR: [#2 Fix eligibility](https://github.com/Klavrin/FarmXpert-backend/pull/2)
- Authors: @kynexi

## 2025-10-30

### Fix Scraper reliability

- Repairs scraper functionality to handle source changes and improve stability. Resulting in fewer scraping errors.
- PR: [#1 Fix scraper](https://github.com/Klavrin/FarmXpert-backend/pull/1)
- Authors: @kynexi

## Notes

- Contributors this period: @kynexi
- Repository language: 100% Python
