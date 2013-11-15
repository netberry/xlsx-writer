<?php

require_once(dirname(__DIR__) . '/src/Netberry/XlsxWriter.php');

echo 'Generating data... ';
$data = array();
for ($i = 0; $i < 10000; $i++) {
    $data[] = array(
        'name' => 'Jack',
        'middle_name' => 'M.',
        'surname' => 'Daniels',
        'age' => rand(25,45),
        'url' => 'http://jack.com',
        'visitors' => rand(1,1000000),
        'comment' => md5(rand(1,1000000) . microtime(true)),
        'color' => 'brown',
        'color_hex' => '#' . rand(10,99) . rand(10,99) . rand(10,99),
    );
    $data[] = array(
        'name' => 'William',
        'middle_name' => '<CoolGuy"&>',
        'surname' => 'Smith',
        'age' => rand(25,45),
        'url' => 'http://jack.com',
        'visitors' => rand(1,1000000),
        'comment' => '\\</t> ' . md5(rand(1,1000000) . microtime(true)),
        // "color" is skipped
        'color_hex' => '#' . rand(10,99) . rand(10,99) . rand(10,99),
        'will_be_not_included' => 'other',
    );
}

echo 'Generating Excel file... ';
$time = microtime(true);
\Netberry\XlsxWriter::write('test.xlsx', $data);
echo 'Excel was generated in ' . round(microtime(true) - $time, 3) . ' seconds. Total rows: ' . count($data);
