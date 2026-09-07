# cbGPS

VB6 GPS map viewer (`cbGPS.exe` / `prjGPS`) that reads NMEA `$GPRMC` sentences over `MSCOMM32.OCX`, parses lat/long, velocity, and direction, and plots them on a zoomable map (`GPSMap` / `frmGPS`). COM port and speed units come from `gpsConfig.ini`. Open `prjGPS.vbp` in the VB6 IDE.

**Source last updated:** 2026-08-27 · **Language:** VB6 · **Target:** VB6 Win32 · **Output:** WinForms exe

## Solution structure

| Project | Language | Type | Purpose |
|---------|----------|------|---------|
| `prjGPS` (`prjGPS.vbp`) | VB6 | WinForms exe | prjGPS |

## How to open

Open the `.vbp` in Visual Basic 6.0 IDE:
- `prjGPS.vbp`

## Requirements

- Visual Basic 6.0 IDE
- Registered OCX/DLL dependencies referenced by the `.vbp` (may need to be installed separately):
  - `MSCOMM32.OCX`

## Attribution and provenance

Working copy from Dave Robinson's OneDrive Historical Dev folder `VB/cbGPS`.

## License

MIT © 2026 VaderConsulting for Dave Robinson's code. See `LICENSE`.
