# Budget optimizer: the planner's Apps Script

`gas.js` is the Google Apps Script web app that reads a client's budget planner (`Pacing & Optimization` tab) and writes suggested weightings back. Since 14 September 2026 its only caller is the Budgets tool in `team-revenu/tools` (`api/budgets/_planner.js`). `index.html` is the retired intake page and redirects to `tools.revenuagency.io`.

Script project in Drive: **Budget optimizer** (`1eySaRtXpjCU_KVos1NdlhqybFq5XTK0vSDVLXmGJHxrQFuE5xvROsMHj`). Deployed URL: `https://script.google.com/macros/s/AKfycbyKod7AuFlgOjIVrwvO1xJhG1fCafvktzGuaJOqUGFqxf-lbGObJcM0rL8j3xRDyHwT/exec` (env `BUDGET_GAS_URL` on the tools project).

## Contract v2 (17 September 2026)

```
POST {action:"read",  token?, sheetId|sheetUrl, tabName?}
  -> {success, sheetId, tabName, headerRow, columnIndices,
      budgetGroups:{<group>:{campaigns:{<campaign>:{currentWeighting, engine, rowIndex}}}},
      rows:[{rowIndex, group, campaign, channel, engine, currentWeighting}], warnings}
POST {action:"write", token?, sheetId|sheetUrl, tabName?, results:[{rowIndex?, budgetGroup, campaign, cost, clicks, cpc, leads, cpl, grossPipeline, qualifiedPipeline, suggestedWeighting, lostIsRank, lostIsBudget}]}
  -> {success, rowsWritten, total, warnings}
GET -> {success:false, error:"POST only"}
```

- `rows` is every campaign row in sheet order. `budgetGroups` is the v1 shape and still collapses a repeated name inside a group; the caller builds groups from `rows`.
- `write` uses `rowIndex` when present, and only if that row's campaign cell still normalises to `campaign`; otherwise it warns and skips the row. Without `rowIndex` it matches `group|campaign`, both sides normalised (case, whitespace, non-breaking spaces).
- `token` is checked against Script Property `BUDGET_TOKEN`. Unset property: no check. Wrong token: `{success:false, error:"unauthorized"}`.
- `tabName` given: that tab must exist. Not given: a tab named like "pacing" or "budget" with "campaign" in its first ten rows, else an error. Never the first sheet.
- `suggestedWeighting` is written as `value / 100` because the column is percent-formatted; weightings are read from display values, so `100%` is read as 100.

## Deploying a new version

1. Open the script project, replace the contents of the one `.gs` file with `gas.js`, save.
2. Project Settings, Script Properties: add `BUDGET_TOKEN` = 32 random bytes as hex (`openssl rand -hex 32`).
3. Deploy, Manage deployments, edit the existing deployment, Version: New version, Deploy. The URL does not change.
4. On the tools Vercel project set `BUDGET_GAS_TOKEN` to the same value (Production and Preview), redeploy.
5. Record the deployment version number here: **v? (date)**.

## The six checks

`T` is the template copy planner `1IJCj0MgkWyWdYLktH1yo2JgZTDuRBSWRh1zd5LPtYv0`, `U` the deployed URL, `K` the token.

```sh
# 1. GET is refused
curl -sL "$U" | jq .                                      # {"success":false,"error":"POST only"}
# 2. wrong token is refused
curl -sL -X POST "$U" -H 'content-type: application/json' -d '{"action":"read","sheetId":"'$T'","token":"nope"}' | jq .error   # "unauthorized"
# 3. read returns rows, tabName and columnIndices
curl -sL -X POST "$U" -H 'content-type: application/json' -d '{"action":"read","sheetId":"'$T'","token":"'$K'"}' | jq '{tabName, headerRow, rows: (.rows|length), groups: (.budgetGroups|keys)}'
# 4. a wrong tab name is an error, not the first sheet
curl -sL -X POST "$U" -H 'content-type: application/json' -d '{"action":"read","sheetId":"'$T'","tabName":"Nope","token":"'$K'"}' | jq .error
# 5. a write by row number lands (pick a real rowIndex and its campaign from check 3)
curl -sL -X POST "$U" -H 'content-type: application/json' -d '{"action":"write","sheetId":"'$T'","token":"'$K'","results":[{"rowIndex":ROW,"budgetGroup":"GROUP","campaign":"CAMPAIGN","suggestedWeighting":12.34}]}' | jq .
# 6. a write to a row that holds another campaign is refused and changes nothing
curl -sL -X POST "$U" -H 'content-type: application/json' -d '{"action":"write","sheetId":"'$T'","token":"'$K'","results":[{"rowIndex":ROW,"budgetGroup":"GROUP","campaign":"Not this one","suggestedWeighting":99}]}' | jq '{rowsWritten, warnings}'   # rowsWritten 0, one warning
```

## History

- 14 Feb 2026: first deployment, composite key writes, batch column runs.
- 3e3a29c: GP/QP columns and engine in the payload (never deployed to Drive).
- 17 Sep 2026: contract v2 (rows, rowIndex writes, token, POST only, display-value weightings). E9 WP-B12b in `tools/BACKLOG.md`.
