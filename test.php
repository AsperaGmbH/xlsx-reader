<?php

use Aspera\Spreadsheet\XLSX\Reader as XLSXReader;
use Aspera\Spreadsheet\XLSX\Worksheet;

if (PHP_SAPI !== 'cli') {
    header('Content-Type: text/plain');
}

if (isset($argv[1])) {
    $filepath = $argv[1];
} elseif (isset($_GET['file'])) {
    $filepath = $_GET['file'];
} else {
    if (PHP_SAPI === 'cli') {
        echo 'Please specify filename as the first argument' . PHP_EOL;
    } else {
        echo 'Please specify filename as a HTTP GET parameter "File", e.g., "/test.php?file=test.xlsx"';
    }

    exit;
}

require('lib/Reader.php');

date_default_timezone_set('UTC');

$StartMem = memory_get_usage();
echo '---------------------------------' . PHP_EOL;
echo 'Starting memory: ' . $StartMem . PHP_EOL;
echo '---------------------------------' . PHP_EOL;

try {
    // set options for initialization
    $reader_options = array(
        'SkipEmptyCells' => true,
        'TempDir'        => sys_get_temp_dir()
    );

    $spreadsheet = new XLSXReader($reader_options);
    $spreadsheet->open($filepath);
    $base_mem = memory_get_usage();

    $sheets = $spreadsheet->getSheets();

    echo '---------------------------------' . PHP_EOL;
    echo 'Spreadsheets:' . PHP_EOL;
    print_r($sheets);
    echo '---------------------------------' . PHP_EOL;
    echo '---------------------------------' . PHP_EOL;

    /** @var Worksheet $sheet_data */
    foreach ($sheets as $Index => $sheet_data) {
        $name = $sheet_data->getName();
        echo '---------------------------------' . PHP_EOL;
        echo '*** Sheet ' . $name . ' ***' . PHP_EOL;
        echo '---------------------------------' . PHP_EOL;

        $Time = microtime(true);

        $spreadsheet->changeSheet($Index);

        foreach ($spreadsheet as $key => $row) {
            echo $key . ': ';

            if ($row) {
                print_r($row);
            } else {
                var_dump($row);
            }

            $current_mem = memory_get_usage();

            echo 'Memory: ' . ($current_mem - $base_mem) . ' current, ' . $current_mem . ' base' . PHP_EOL;
            echo '---------------------------------' . PHP_EOL;

            if ($key && ($key % 500 === 0)) {
                echo '---------------------------------' . PHP_EOL;
                echo 'Time: ' . (microtime(true) - $Time);
                echo '---------------------------------' . PHP_EOL;
            }
        }

        echo PHP_EOL . '---------------------------------' . PHP_EOL;
        echo 'Time: ' . (microtime(true) - $Time);
        echo PHP_EOL;

        echo '---------------------------------' . PHP_EOL;
        echo '*** End of sheet ' . $name . ' ***' . PHP_EOL;
        echo '---------------------------------' . PHP_EOL;
    }

} catch (Exception $e) {
    echo $e->getMessage();
}
