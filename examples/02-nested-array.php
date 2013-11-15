<?php

require_once(dirname(__DIR__) . '/src/Netberry/XlsxWriter.php');

echo 'Generating data... ';
$data = array();
for ($i = 0; $i < 10000; $i++) {
    $data[] = array(
        'name' => 'Jack',
        'colors' => array('black', 'blue', 'red'),
    );
    $data[] = array(
        'name' => 'Jasmine',
        'colors' => array('green', 'yellow', 'red' => array('pink', 'magenta'))
    );
}

echo 'Generating Excel file... ';
$time = microtime(true);
\Netberry\XlsxWriter::write('test.xlsx', $data);
echo 'Excel was generated in ' . round(microtime(true) - $time, 3) . ' seconds. Total rows: ' . count($data);
