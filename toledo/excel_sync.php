<?php

define('INVENTORY_SYNC_TARGET', __DIR__ . DIRECTORY_SEPARATOR . 'inventory_realtime_export_20260325_104837_excel_app.xlsx');

function syncInventoryWorkbook($conn = null, $targetFile = null)
{
    if ($targetFile === null) {
        $targetFile = INVENTORY_SYNC_TARGET;
    }

    $script = __DIR__ . DIRECTORY_SEPARATOR . 'sync_excel_app.ps1';

    if (!file_exists($script)) {
        return false;
    }

    $command = 'powershell -NoProfile -ExecutionPolicy Bypass -File '
        . escapeshellarg($script)
        . ' -OutputPath '
        . escapeshellarg($targetFile);

    exec($command, $output, $exitCode);

    return $exitCode === 0 && file_exists($targetFile);
}

function syncInventoryMessage($baseMessage, $syncOk)
{
    if ($syncOk) {
        return urlencode($baseMessage);
    }

    return urlencode($baseMessage . ' Excel sync failed. Close the workbook if it is open, then save again.');
}
