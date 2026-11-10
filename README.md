# Sootcase for Google Sheets

Freight emissions in your spreadsheet: a `=SOOTCASE()` cell function and a menu that calculates a whole sheet of shipments. Calculations come from [Sootcase](https://sootcase.org) and follow the GLEC Framework v3.3 and ISO 14083.

```
=SOOTCASE("Copenhagen", "Hamburg", 12, "road")          → kg CO2e, well-to-wheel
=SOOTCASE("CPH", "FRA", 0.8, "air")                     → by IATA code
=SOOTCASE("DEHAM", "NLRTM", 20, "sea", "container_ship") → by UN/LOCODE
=SOOTCASE(A2:A500, B2:B500, C2:C500, "road")             → a whole column in one request
```

## Install

1. Open your spreadsheet, then **Extensions → Apps Script**.
2. Delete what's in `Code.gs` and paste in [`src/Code.js`](src/Code.js).
3. Click **Project Settings** (gear icon), tick **Show "appsscript.json" manifest file**, open `appsscript.json` and paste in [`src/appsscript.json`](src/appsscript.json).
4. Save, close the Apps Script tab and reload the spreadsheet. A **Sootcase** menu appears.
5. **Sootcase → Set my API key…** and paste the key from [sootcase.org](https://sootcase.org) (sign in with your email). The first time, Google asks you to allow the script to reach sootcase.org.

With [clasp](https://github.com/google/clasp) instead: `cp .clasp.json.example .clasp.json`, put in your script ID, then `npx @google/clasp push`.

## The cell function

`SOOTCASE(from, to, tonnes, mode, [vehicle], [distance_km], [value])`

| Argument | |
|---|---|
| `from`, `to` | an upper-case UN/LOCODE (`NLRTM`) or IATA code (`CPH`), a place name (`Rotterdam`, the best match is used), or `"lat,lon"` |
| `tonnes` | cargo weight in tonnes |
| `mode` | `road`, `rail`, `sea`, `air` or `inland_waterway` (also `truck`, `train`, `ship`, `plane`, `barge`) |
| `vehicle` | optional; default `hgv`, `freight_train`, `container_ship`, `air_freight`, `barge`. Any Sootcase vehicle type works, e.g. `hgv_artic`, `freight_train_electric`, `container_ship_deep_sea` |
| `distance_km` | optional; otherwise Sootcase routes the leg (road, rail, sea and waterway networks; great-circle for air) |
| `value` | optional: `wtw` (default), `ttw`, `wtt`, `tonne_km`, `intensity` (g CO2e per tonne-km), `distance_km` |

A range is sent in requests of 500 rows, all at once (a cell function gets 30 seconds), and every answer is cached for 6 hours, so recalculating the sheet doesn't call the API again. A bad row shows `Error: …` in its cell; the rest still calculate.

Sootcase adds road legs where they're needed (to a port or airport, for example), so `distance_km` as a `value` is the whole trip.

## A whole sheet of shipments

For multi-leg shipments, hubs, measured fuel, loads or your own factors, use Sootcase's CSV columns, one row per leg or hub stop:

1. **Sootcase → Insert shipments template** gives a sheet with the columns and a few examples.
2. Fill it in. Every column of [Sootcase's CSV](https://github.com/majidkabir/sootcase#inputs) works: `fuel`, `fuel_amount`, `load_tonnes`, `own_wtw_g_per_tkm`, `hub_type` and so on.
3. **Sootcase → Calculate shipments on this sheet** (up to 10,000 shipments).

It writes three sheets:

- **Sootcase results**: per shipment, emissions (TTW, WTT, WTW, hubs), tonne-km, intensity, share of primary data, road legs Sootcase added, warnings (things that may make a result too low), or why it failed.
- **Sootcase legs**: per leg and hub, distance and how it was found, km per country (routed legs), DAF, method, data type and the factor's source, so every number can be checked.
- **Sootcase summary**: the ISO 14083-style report: totals, per mode, hubs, data quality with the default value sources, and notes.

## API key and server

- **Set my API key…** stores the key for you only.
- If `=SOOTCASE()` cells say the key is missing, Google isn't giving cell functions your own settings. Use **Set API key for this spreadsheet…** instead. Anyone who can edit the spreadsheet's script can read that key.
- **Use another Sootcase server…** points the add-in at your own `sootcase serve` (e.g. `http://your-host:8080`).

## Development

```sh
go install github.com/majidkabir/sootcase/cmd/sootcase@latest
sootcase serve -addr :8080 &
npm test      # runs src/Code.js in Node against that server
```

The tests use small stand-ins for the Apps Script services (`test/harness.js`) and a real Sootcase server, with given distances, so no routing data is needed.

## License

MIT. The calculations are Sootcase's ([PolyForm Noncommercial](https://github.com/majidkabir/sootcase/blob/main/LICENSE)), with GLEC Framework default values used with the permission of Smart Freight Centre. Results are calculations for your own reporting; Sootcase isn't certified against ISO 14083.
