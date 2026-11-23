# Home Assistant names and aliases

`Set_HAAliases.ps1` sets entity names, entity aliases (used by Assist, Google, and Alexa), and device names through the Home Assistant WebSocket API. Changes are live. No restart is needed, and entity IDs don't change, so automations keep working.

## Workflow

1. **Back up Home Assistant.** Settings → System → Backups → Create backup. The script writes a rollback CSV, but a full backup is the safety net.
2. **Export the current state.**
   ```powershell
   .\Set_HAAliases.ps1 -Export
   ```
   This writes `aliases_export_<timestamp>.csv` with every entity and device. `NewName` and `Aliases` are pre-filled with the current values, so you edit them in place. `DisplayedName`, `Area`, and `Device` are for reference only. You are prompted for a long-lived access token (Profile → Security → Long-lived access tokens). Set `HA_ACCESS_TOKEN` to skip the prompt.
3. **Build `aliases.csv`.** Save the export as `aliases.csv`, then edit `NewName` and `Aliases` on the rows you want to change. Rows you don't change can stay. A value that matches the current one is a no-op.
4. **Preview.** Run without `-Apply`. Nothing is written.
   ```powershell
   .\Set_HAAliases.ps1
   ```
5. **Apply.** A rollback CSV is written to `Backups\` first, then the changes are made.
   ```powershell
   .\Set_HAAliases.ps1 -Apply
   ```
6. **Undo, if needed.** Apply the rollback file as a mapping file:
   ```powershell
   .\Set_HAAliases.ps1 -MappingFile ".\Backups\aliases_rollback_<timestamp>.csv" -Apply
   ```

## CSV format

| Column | Meaning |
|---|---|
| `Type` | `entity` or `device` |
| `Target` | `entity_id` for entities (for example `light.kitchen_ceiling`). Device id or device name for devices. |
| `NewName` | New name. Pre-filled with the current custom name. Empty, or the same as the current value, = unchanged. `<clear>` = remove the custom name and use the default. |
| `Aliases` | Entities only. Semicolon-separated list, for example `kitchen light;ceiling lights`. Pre-filled with the current aliases. Empty, or the same as the current value, = unchanged. `<name>` = the entity's friendly name. `<clear>` = remove all aliases. |
| `DisplayedName` | Read-only. The name Home Assistant shows now. Ignored on apply. |
| `Area`, `Device` | Read-only. Ignored on apply. |

Because an empty cell means "unchanged", use `<clear>` to remove a name or all aliases.

The first alias is what Google and Alexa use as the device name. Put the name you say most often first.

Device aliases are not supported by Home Assistant. Device rows with `Aliases` are ignored with a warning.

The script validates the whole file before it connects. It rejects bad types, invalid entity IDs, and duplicate rows. It also rejects device names that match more than one device, so use the device id in that case.

## Notes

- `Backups\` and `aliases_export_*.csv` are git-ignored. They contain your device and entity names.
- The access token is never written to disk. Don't pass it on the command line, because that puts it in shell history.
- Renaming an entity does not change its entity ID, so automations, scripts, and dashboards keep working. Other places that show the old name, such as dashboard cards, update automatically.
