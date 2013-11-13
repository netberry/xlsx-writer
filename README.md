xlsx-writer
===========

xlsx Writer - write real Excel xlsx in PHP

## Installation

Install with [Composer](http://getcomposer.org).

```json
{
	"require": {
		"netberry/xlsxwriter": "0.0.1"
	}
}
```

## Usage

```php
require 'vendor/autoload.php';

\Netberry\XlsxWriter::write('test.xlsx', array(
    array('name' => 'Brad', 'surname' => 'Pitt'),
    array('name' => 'Leonardo', 'surname' => 'DiCaprio'),
));
```

## Requirements
 * PHP version 5.2.0 or higher
 * PHP extension php_zip enabled

## License
MIT