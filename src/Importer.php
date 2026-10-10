<?php

/**
 * (c) Joffrey Demetz <joffrey.demetz@gmail.com>
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace JDZ\Importer;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;

/**
 * Spreadsheet importer — reads the active worksheet into an array of
 * field-keyed rows, coercing each value to its declared column type.
 *
 * @author Joffrey Demetz <joffrey.demetz@gmail.com>
 */
class Importer
{
  /**
   * Parse a spreadsheet file (format sniffed by PhpSpreadsheet).
   *
   * The first row is the header; columns come from getColumns() when the
   * subclass declares them, otherwise from the header labels themselves.
   */
  public function parse(string $tmpfile): array
  {
    $spreadsheet = IOFactory::load($tmpfile);
    $sheet = $spreadsheet->getActiveSheet();
    $data = $sheet->toArray(null, true, true, true);

    $rows = [];

    if (empty($data)) {
      return $rows;
    }

    // rows stay keyed by their sheet row number (array_shift() would renumber them)
    $_columns = reset($data);
    unset($data[key($data)]);

    if (!($columns = $this->getColumns())) {
      $columns = [];
      foreach ($_columns as $column) {
        $columns[] = (object)['field' => $column];
      }
    }

    foreach ($columns as $column) {
      if (!isset($column->default)) {
        $column->default = '';
      }
      if (!isset($column->type)) {
        $column->type = 'string';
      }
    }

    foreach ($data as $rowNumber => $dataRow) {
      $row = [];

      for ($c = 0, $cn = count($columns); $c < $cn; $c++) {
        $letter = Coordinate::stringFromColumnIndex($c + 1);

        $value = $dataRow[$letter] ?? $columns[$c]->default;
        $value = trim((string)$value);

        // a date cell is read from its serial: its formatted value follows the cell's
        // number format (Excel's built-in date format renders 03-05-24, month first)
        if ('date' === $columns[$c]->type && isset($dataRow[$letter])) {
          $cell = $sheet->getCell($letter . $rowNumber);
          if (is_numeric($cell->getValue()) && Date::isDateTime($cell)) {
            $value = Date::excelToDateTimeObject((float)$cell->getValue())->format('Y-m-d');
          }
        }

        switch ($columns[$c]->type) {
          case 'int':
            $value = (int)$value;
            break;

          case 'bool':
            $value = 1 === intval($value) || 'X' === strtoupper($value);
            break;

          case 'url':
          case 'string':
          case 'date':
          default:
            $value = (string)$value;
            break;
        }

        $row[$columns[$c]->field] = $this->cleanValue($value, $columns[$c]);
      }

      $rows[] = $row;
    }

    return $rows;
  }

  /**
   * Declared columns, in sheet order — \stdClass{field, type?, default?}.
   */
  protected function getColumns(): array
  {
    return [];
  }

  protected function cleanValue(mixed $value, \stdClass $column): mixed
  {
    if ('url' === $column->type) {
      if ('' !== $value && !preg_match("/^https?:\/\//", $value)) {
        $value = 'https://' . $value;
      }

      return str_replace('http://', 'https://', $value);
    }

    if ('bool' === $column->type) {
      return 1 === intval($value) ? 1 : 0;
    }

    if ('date' === $column->type) {
      if ('' !== $value) {
        if (preg_match("/^(\d{2})\/(\d{2})\/(\d{4}).*$/", $value, $m)) {
          $value = $m[3] . '-' . $m[2] . '-' . $m[1];
        } elseif (preg_match("/^(\d{2})-(\d{2})-(\d{4}).*$/", $value, $m)) {
          $value = $m[3] . '-' . $m[2] . '-' . $m[1];
        } elseif (preg_match("/^(\d{4})-(\d{2})-(\d{2}).*$/", $value, $m)) {
          $value = $m[1] . '-' . $m[2] . '-' . $m[3];
        } else {
          $value = '';
        }
      }
    }

    return $value;
  }
}
