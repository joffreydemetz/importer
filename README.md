# JDZ Importer

Spreadsheet import (xlsx / csv / ods) built on [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet).

Reads the active worksheet of an uploaded file into an array of field-keyed rows. The first row is the header, and each value is coerced to its column type (string / int / bool / url / date).

```php
use JDZ\Importer\Importer;

$rows = (new Importer())->parse('/tmp/upload.xlsx');
// [ ['name' => 'Doe', 'email' => 'doe@example.com'], … ]
```

Header labels become the field names by default. Subclass `getColumns()` to map the sheet onto your own fields and types:

```php
class MembersImporter extends Importer
{
    protected function getColumns(): array
    {
        return [
            (object)['field' => 'name'],
            (object)['field' => 'active',  'type' => 'bool',  'default' => 0],
            (object)['field' => 'website', 'type' => 'url'],
            (object)['field' => 'joined',  'type' => 'date'],
        ];
    }
}
```

Coercions: `bool` reads `1` or `X` (yields `1`/`0`), `url` forces `https://`, `date` normalizes `dd/mm/yyyy` and `dd-mm-yyyy` to `Y-m-d` (anything unparseable becomes an empty string).

## Install

```
composer require jdz/importer
```

## Tests

```
composer test
```
