# Reclami: customer complaint management for a manufacturing quality office

*Gestione dei reclami clienti per l'ufficio qualità di un'azienda manifatturiera*

**Visual Basic 6** · 2000–2003 · version 7.3.4  
Author: **Massimo Sbarbaro** ([ORCID 0009-0006-8965-9013](https://orcid.org/0009-0006-8965-9013))

## Overview

Reclami is the central application of the quality office of a plastic-film manufacturer. Every customer complaint is linked to the production order (ODL) it refers to and is analysed from two sides: a **technical analysis** (defect, defect family, acceptance or rejection, corrective actions, quality index) and a **commercial analysis** (returned kilograms, credit note, amount credited). Complaint cards written by the sales offices arrive by e-mail and are loaded automatically; order and shipment data are imported from the company ERP.

I designed and programmed this application in 2000–2003. It is published here in 2026 as part of the documented record of my software work, with a permanent DOI on Zenodo.

## Main functions

- Tree view of all complaints by status, opening the technical (`frmATC`) or commercial (`frmMKG`) analysis.
- Built-in POP3 client (Winsock) that reads the e-mails sent by the sales offices, recognised by a subject tag and delimiter strings, and inserts them into the `ReclamoVenditore` table (`frmPop3`).
- Import of order and shipment data from the MFG/PRO ERP: an Attachmate Reflection script runs a query on the host, the result is downloaded by FTP and merged into the `MFG` table (`frmImportMFG`).
- Quality index, defect-family and sales-network statistics, with Crystal Reports and Microsoft Access reports driven by automation (`frmIQ`, `frmRete`, `frmRicerca`, `frmGrafici`).
- Role-based access read from a local `UTENTE.ini` file (technical, commercial, read-only).

## Data

Microsoft Access (Jet) database accessed through DAO 3.6. Main tables: `ATC`, `MKG`, `MFG`, `ReclamoVenditore`, `Famiglia`, `Difetto`, `Linea`, `Causale`, `Parametri`, `Indici`.

## Technology

DAO 3.6, Microsoft Access object library, Crystal Reports 8.5 (RDC and viewer), Winsock, Common Controls, Tab control, Masked Edit.

## Repository contents

| Path | Content |
|---|---|
| `src/` | Visual Basic 6 project (`.vbp`), forms (`.frm` with their binary resources `.frx`) and modules (`.bas`), as listed in the project file. |
| `config-example/` | Templates of the `.ini` configuration files read at start-up, with placeholder values. |

## What is not included

Crystal Reports layouts (`.rpt`), compiled executables, installers, scripts for the host connection and the production databases are **not** included, because they contain operational data of the organisations for which the program was built. Configuration files are published as templates in `config-example/` with placeholder values: server addresses, accounts and passwords have been removed. Names of client organisations and products have been removed from captions and comments; local paths have been normalised.

## Related repositories

- [complaint-card-sales-network-vb6](https://github.com/massimosbarbaro/complaint-card-sales-network-vb6)
- [erp-shipments-to-sales-networks-vb6](https://github.com/massimosbarbaro/erp-shipments-to-sales-networks-vb6)

## How to cite

Use the citation metadata in [`CITATION.cff`](CITATION.cff) (GitHub: *Cite this repository*). Each release is archived on Zenodo with its own DOI.

> Sbarbaro, Massimo. *Reclami: customer complaint management for a manufacturing quality office (Visual Basic 6, 2000–2003)*. Software, version 7.3.4. GitHub: https://github.com/massimosbarbaro/quality-complaints-manager-vb6

## License

Released under the [MIT License](LICENSE). © 2000 Massimo Sbarbaro.
