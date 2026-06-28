# Experimental Simkl Target Plan

Status: first implementation draft  
Reviewed: 2026-06-24  
Proposed branch: `experimental/simkl-target`

Update, 2026-06-28: the experimental branch now uses a separate Windows app
identity from the normal tray release: `PlexTraktSync Tray Experimental`,
`PlexTraktSyncTrayExperimental.exe`, a separate scheduled task, a separate
Start menu shortcut, a separate mutex, a separate Simkl Credential Manager
service, and a separate local runtime folder at
`%LOCALAPPDATA%\PlexTraktSyncTrayExperimental`. It still reads the underlying
PlexTraktSync config and watch log from `%LOCALAPPDATA%\PlexTraktSync\PlexTraktSync`.

## Decision

Simkl is a good candidate for an experimental target. It has an official,
documented API; explicitly supports desktop binaries and media-server plugins;
and accepts episode-level history writes with the original watch timestamp.

The first version should be an opt-in history target on top of the existing
multi-target ledger. It should not replace PlexTraktSync's Trakt scrobbler or
attempt live Simkl scrobbling.

## V1 Scope

- Connect one Simkl account through Simkl's PIN flow.
- Store the access token in Windows Credential Manager through `keyring`.
- Send completed Plex movie and episode events to Simkl.
- Preserve the Plex completion timestamp as `watched_at` in UTC.
- Check the specific movie or episode before writing when local state is not
  already conclusive.
- Record each outcome in the existing `target_attempts` ledger under `simkl`.
- Show concise Simkl connection and last-sync state in the tray menu and log.
- Keep the feature disabled until both a client ID and user token exist.

Not in V1:

- Importing historical Trakt data.
- Ratings. Plex completion events do not provide a trustworthy 1-10 rating.
- Live start, pause, stop, or progress scrobbling.
- Deleting or reconciling Simkl history.
- Rewatches. Simkl's separate rewatch history is Pro/VIP-only and needs an
  explicit product decision.
- Anime-specific numbering beyond Plex's normal season/episode mapping.

## Official API Contract

Base API URL: `https://api.simkl.com`

Every request must include:

- `client_id`, `app-name`, and `app-version` query parameters.
- A descriptive `User-Agent` header.
- `Authorization: Bearer <token>` for user data.

Relevant endpoints:

| Purpose | Method and endpoint | Notes |
| --- | --- | --- |
| Start PIN login | `GET /oauth/pin` | Returns a five-character `user_code`, verification URI, 15-minute expiry, and polling interval. |
| Poll PIN login | `GET /oauth/pin/{user_code}` | Poll at the returned interval, currently five seconds. This is Simkl-specific and not standard RFC 8628 wire format. |
| Check watched state | `POST /sync/watched` | Include `season` and `episode` for an exact episode check. Do not infer episode state from show-level `result` or list status. |
| Add history | `POST /sync/history` | Supports movie and exact episode writes with per-item `watched_at`. |
| Validate account | `POST /users/settings` | Useful once after login to verify the token and cache account type. |

Simkl accepts the IDs already collected in `MediaEvent`, including IMDb, TMDB,
and TVDB. Send every available ID. A TMDB ID is only unique when paired with
the media type, so the request must retain the movie/show container shape.

Example episode write:

```json
{
  "shows": [
    {
      "ids": {
        "imdb": "tt1234567",
        "tmdb": 12345,
        "tvdb": 67890
      },
      "seasons": [
        {
          "number": 2,
          "episodes": [
            {
              "number": 4,
              "watched_at": "2026-06-19T03:15:00Z"
            }
          ]
        }
      ]
    }
  ]
}
```

## Authentication Design

Register PlexTraktSync Tray as a public Simkl application. The client ID may be
distributed with the app; no client secret should be embedded in the source or
release binary.

Suggested flow:

1. Add `Connect Simkl` to the experimental tray menu.
2. Request a PIN and open `https://simkl.com/pin` in the user's browser.
3. Put the short code in both a notification and the tray status line.
4. Poll in one background thread at the server-provided interval.
5. Stop on success, expiry, shutdown, or an unexpected replacement-code
   response.
6. Store the token under service `PlexTraktSyncTrayExperimental.Simkl`, username `oauth`.
7. Call `/users/settings` once to validate the token and cache the account ID,
   username, and plan type without logging the token.
8. Offer `Disconnect Simkl`, which removes only the stored Simkl credential and
   cached account metadata.

Environment variables can remain a developer-only override:

- `SIMKL_CLIENT_ID`
- `SIMKL_TOKEN`

## Target Design

Add `SimklTarget(SyncTarget)` and register it with `TargetDispatcher`.

`is_configured()` should require a client ID and access token. `applies_to()`
should accept completed movies and normal episodes. `sync()` should:

1. Reject events with no usable external IDs as `blocked`.
2. Reject episodes missing season or episode numbers as `blocked`.
3. Build an exact `/sync/watched` lookup using the IDs and, for episodes, the
   season and episode numbers.
4. Return `already_present` only when that exact item is reported watched.
5. Submit one `/sync/history` write with the event's UTC timestamp.
6. Treat HTTP success as provisional. Confirm that all `not_found` arrays are
   empty and that the expected item count is represented in the response.
7. Return `synced`, `already_present`, `blocked`, or `failed_retryable` for the
   ledger.

The local ledger remains the primary fast idempotency check. The API lookup is
the defense against entries created outside this tray, a rebuilt ledger, or an
earlier ambiguous response.

## Rate Limits and Retries

Simkl documents limits of 10 GET requests per second and one POST request per
second per client ID and per user token. Sync endpoints should be sequential.

- Put all Simkl POSTs behind one process-wide lock and enforce at least one
  second between write attempts.
- Do not parallelize duplicate checks and writes for multiple events.
- Retry `429`, `500`, `502`, and `503` with exponential backoff, capped at 60
  seconds and five attempts.
- Honor `Retry-After` when present.
- Treat `400 RATE_LIMIT` as retryable because Simkl may use it for a per-user
  lock collision.
- Treat `401` as an auth failure and stop retrying until the user reconnects.
- Treat `412 client_id_failed` as a configuration/release error.
- Never poll `/sync/all-items` for this event-driven target.

## Configuration and UI

Keep non-secret settings in a small JSON file under the existing local app-data
directory. Suggested fields:

```json
{
  "enabled": false,
  "client_id": "",
  "account_id": null,
  "username": null,
  "account_type": null
}
```

The public experimental build can bundle the project's registered client ID.
Developer builds can override it with `SIMKL_CLIENT_ID`. The access token must
never be written to JSON, logs, notifications, SQLite request summaries, crash
messages, or release artifacts.

Suggested menu state:

- `Simkl: not connected`
- `Simkl: waiting for PIN ABCDE`
- `Simkl: connected as <username>`
- `Simkl: auth failed`
- `Connect Simkl` / `Reconnect Simkl`
- `Disconnect Simkl`
- `Open Simkl`

## Public Release Gate

Before publishing a binary for other people:

- Register the project in Simkl developer settings and use that app's client
  ID, name, version, and `User-Agent` consistently.
- Link back to Simkl from the tray menu and README.
- Keep requests sequential and within the documented limits; do not ship a
  polling loop that repeatedly downloads a user's full library.
- Confirm the project still qualifies under Simkl's published licensing rule:
  non-commercial projects and commercial projects below $150/month are free;
  projects at or above that threshold need a commercial license.
- Describe the integration as official-API-based but not affiliated with or
  endorsed by Simkl.
- Publish it as experimental until duplicate handling and token recovery have
  been exercised by more than one account.

## Testing

Unit tests:

- Movie and episode request payloads include all available IDs.
- `watched_at` is converted to UTC without losing the original instant.
- Exact episode lookup includes both season and episode.
- Show-level `result: true` cannot suppress an episode write unless the exact
  episode lookup is true.
- `not_found` in a 201 response is not recorded as success.
- Empty or zero-count success responses are not accepted blindly.
- PIN pending, success, expiry, and replacement-code responses stop correctly.
- Tokens and authorization headers are redacted from all summaries and errors.
- Retry classification covers `400 RATE_LIMIT`, `401`, `412`, `429`, and 5xx.
- The one-POST-per-second limiter works across concurrent target calls.

Manual validation on a dedicated Simkl test account:

1. Connect through PIN and restart the tray to prove credential persistence.
2. Play one movie and one episode to completion.
3. Verify title, season/episode, and timestamp on Simkl.
4. Replay the same source log event and verify no duplicate history entry.
5. Pre-create an episode on Simkl, then process it through the tray and verify
   `already_present`.
6. Test an unknown ID and confirm the ledger records a non-terminal failure.
7. Disconnect and confirm no further Simkl requests are attempted.

## Implementation Sequence

1. Create `experimental/simkl-target` from the multi-target ledger branch after
   the current uncommitted work is resolved.
2. Add configuration, Credential Manager storage, and a redacting HTTP helper.
3. Implement and test PIN authentication independently of target sync.
4. Implement exact watched lookup and history payload builders with unit tests.
5. Add `SimklTarget`, sequential rate limiting, retries, and ledger outcomes.
6. Add tray status/actions and structured logs.
7. Validate against a dedicated account before asking anyone else to test it.
8. Publish only as a clearly labeled experimental prerelease. Do not merge to
   `main` until auth recovery, duplicate handling, and timestamp correctness
   have survived real use.

## Open Decisions

- Register one project-owned Simkl client ID or require testers to supply their
  own during the first private experiment.
- Sync movies as well as episodes in V1. The API supports both; enabling both is
  the more coherent default, but it expands the behavior beyond the current
  Serializd episode-only target.
- Decide later whether Pro/VIP rewatches should be opt-in. They must never be
  inferred merely because the same title appears again.
- Decide whether Simkl support belongs in this tray long term or in a separate
  multi-service sync project after the experiment proves useful.

## Sources

- [Simkl API overview](https://api.simkl.org/)
- [Authentication options](https://api.simkl.org/authentication)
- [PIN flow](https://api.simkl.org/api-reference/pin)
- [Add to history](https://api.simkl.org/api-reference/simkl/add-to-history)
- [Look up watched state](https://api.simkl.org/api-reference/simkl/get-watched)
- [Rate limits](https://api.simkl.org/resources/rate-limits)
- [API rules and licensing](https://api.simkl.org/api-rules)
- [OpenAPI 3.1 schema](https://api.simkl.org/openapi.json)
