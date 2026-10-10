<?php

/**
 * (c) Joffrey Demetz <joffrey.demetz@gmail.com>
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace JDZ\Importer\Tests;

use JDZ\Importer\Importer;
use PHPUnit\Framework\Attributes\CoversClass;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PHPUnit\Framework\Attributes\DataProvider;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PHPUnit\Framework\TestCase;

#[CoversClass(Importer::class)]
class ImporterTest extends TestCase
{
    /** @var string[] */
    private array $tmpFiles = [];

    protected function tearDown(): void
    {
        foreach ($this->tmpFiles as $path) {
            if (is_file($path)) {
                unlink($path);
            }
        }
        $this->tmpFiles = [];
    }

    /**
     * Write a sheet (array of rows, first row = header) to a temp file.
     */
    private function sheetFile(array $rows, string $format = 'Xlsx'): string
    {
        $spreadsheet = new Spreadsheet();
        $spreadsheet->getActiveSheet()->fromArray($rows, null, 'A1');

        $path = $this->tmpPath(strtolower($format));
        IOFactory::createWriter($spreadsheet, $format)->save($path);

        return $path;
    }

    /**
     * A fresh temp path with that extension, removed in tearDown (tempnam() left its own file behind).
     */
    private function tmpPath(string $extension): string
    {
        $path = sys_get_temp_dir() . '/jdz-importer-' . uniqid('', true) . '.' . $extension;
        $this->tmpFiles[] = $path;

        return $path;
    }

    public function testHeaderLabelsBecomeFields(): void
    {
        $path = $this->sheetFile([
            ['name', 'email'],
            ['Doe', 'doe@example.com'],
            ['Roe', 'roe@example.com'],
        ]);

        $rows = (new Importer())->parse($path);

        $this->assertCount(2, $rows);
        $this->assertSame(['name' => 'Doe', 'email' => 'doe@example.com'], $rows[0]);
        $this->assertSame(['name' => 'Roe', 'email' => 'roe@example.com'], $rows[1]);
    }

    public function testValuesAreTrimmed(): void
    {
        $path = $this->sheetFile([
            ['name'],
            ['  Doe  '],
        ]);

        $rows = (new Importer())->parse($path);

        $this->assertSame('Doe', $rows[0]['name']);
    }

    public function testDeclaredColumnsDriveFieldsAndTypes(): void
    {
        $path = $this->sheetFile([
            ['Nom', 'Actif', 'Site', 'Age'],
            ['Doe', 'X', 'example.com', '42'],
        ]);

        $importer = new class extends Importer {
            protected function getColumns(): array
            {
                return [
                    (object)['field' => 'name'],
                    (object)['field' => 'active', 'type' => 'bool'],
                    (object)['field' => 'website', 'type' => 'url'],
                    (object)['field' => 'age', 'type' => 'int'],
                ];
            }
        };

        $rows = $importer->parse($path);

        $this->assertSame('Doe', $rows[0]['name']);
        $this->assertSame(1, $rows[0]['active']);
        $this->assertSame('https://example.com', $rows[0]['website']);
        $this->assertSame(42, $rows[0]['age']);
    }

    public function testBoolReadsOneOrX(): void
    {
        $path = $this->sheetFile([
            ['flag'],
            ['X'],
            ['1'],
            ['x'],
            [''],
            ['nope'],
        ]);

        $importer = new class extends Importer {
            protected function getColumns(): array
            {
                return [(object)['field' => 'flag', 'type' => 'bool']];
            }
        };

        $rows = $importer->parse($path);

        $this->assertSame([1, 1, 1, 0, 0], array_column($rows, 'flag'));
    }

    public function testUrlForcesHttps(): void
    {
        // the id column keeps the blank-site row from being trimmed as empty
        $path = $this->sheetFile([
            ['id', 'site'],
            ['1', 'http://example.com'],
            ['2', 'https://example.com'],
            ['3', 'example.com'],
            ['4', ''],
        ]);

        $importer = new class extends Importer {
            protected function getColumns(): array
            {
                return [
                    (object)['field' => 'id'],
                    (object)['field' => 'site', 'type' => 'url'],
                ];
            }
        };

        $rows = $importer->parse($path);

        $this->assertSame([
            'https://example.com',
            'https://example.com',
            'https://example.com',
            '',
        ], array_column($rows, 'site'));
    }

    public function testDateNormalization(): void
    {
        // the id column keeps the blank-day row from being trimmed as empty
        $path = $this->sheetFile([
            ['id', 'day'],
            ['1', '31/12/2024'],
            ['2', '31-12-2024'],
            ['3', '2024-12-31'],
            ['4', 'not a date'],
            ['5', ''],
        ]);

        $importer = new class extends Importer {
            protected function getColumns(): array
            {
                return [
                    (object)['field' => 'id'],
                    (object)['field' => 'day', 'type' => 'date'],
                ];
            }
        };

        $rows = $importer->parse($path);

        $this->assertSame([
            '2024-12-31',
            '2024-12-31',
            '2024-12-31',
            '',
            '',
        ], array_column($rows, 'day'));
    }

    public function testMissingCellFallsBackToDefault(): void
    {
        // second row only fills the first column
        $path = $this->sheetFile([
            ['name', 'role'],
            ['Doe'],
        ]);

        $importer = new class extends Importer {
            protected function getColumns(): array
            {
                return [
                    (object)['field' => 'name'],
                    (object)['field' => 'role', 'default' => 'member'],
                ];
            }
        };

        $rows = $importer->parse($path);

        $this->assertSame('Doe', $rows[0]['name']);
        $this->assertSame('member', $rows[0]['role']);
    }

    public function testMoreThan26Columns(): void
    {
        $header = [];
        $values = [];
        for ($i = 1; $i <= 30; $i++) {
            $header[] = 'col' . $i;
            $values[] = 'v' . $i;
        }

        $path = $this->sheetFile([$header, $values]);

        $rows = (new Importer())->parse($path);

        $this->assertCount(30, $rows[0]);
        $this->assertSame('v27', $rows[0]['col27']);
        $this->assertSame('v30', $rows[0]['col30']);
    }

    public function testCsvIsParsedToo(): void
    {
        $path = $this->sheetFile([
            ['name', 'email'],
            ['Doe', 'doe@example.com'],
        ], 'Csv');

        $rows = (new Importer())->parse($path);

        $this->assertSame(['name' => 'Doe', 'email' => 'doe@example.com'], $rows[0]);
    }

    public function testEmptySheetReturnsNoRows(): void
    {
        $path = $this->sheetFile([['name', 'email']]);

        $this->assertSame([], (new Importer())->parse($path));
    }


    private static function dayImporter(): Importer
    {
        return new class extends Importer {
            protected function getColumns(): array
            {
                return [
                    (object)['field' => 'id'],
                    (object)['field' => 'day', 'type' => 'date'],
                ];
            }
        };
    }

    public static function excelDateFormats(): array
    {
        return [
            'dd/mm/yyyy' => [NumberFormat::FORMAT_DATE_DDMMYYYY],
            'yyyy-mm-dd' => [NumberFormat::FORMAT_DATE_YYYYMMDD],
            'built-in 14 (mm-dd-yy), what Excel saves for a typed date' => [NumberFormat::FORMAT_DATE_XLSX14],
            'with a time' => [NumberFormat::FORMAT_DATE_DATETIME],
        ];
    }

    /**
     * 5 March 2024 (day and month both <= 12, so a swap would go unnoticed otherwise).
     */
    #[DataProvider('excelDateFormats')]
    public function testAnExcelDateCellIsReadAsAnIsoDate(string $formatCode): void
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->fromArray([['id', 'day'], ['1', 45356.5]], null, 'A1');
        $sheet->getStyle('B2')->getNumberFormat()->setFormatCode($formatCode);

        $path = $this->tmpPath('xlsx');
        IOFactory::createWriter($spreadsheet, 'Xlsx')->save($path);

        $this->assertSame('2024-03-05', self::dayImporter()->parse($path)[0]['day']);
    }

    public function testATypedDateKeepsTheDayOnlyWhateverItsTime(): void
    {
        $path = $this->sheetFile([
            ['id', 'day'],
            ['1', '05/03/2024 10:30'],
            ['2', '05-03-2024 10:30'],
            ['3', '2024-03-05 10:30:00'],
        ]);

        $this->assertSame(
            ['2024-03-05', '2024-03-05', '2024-03-05'],
            array_column(self::dayImporter()->parse($path), 'day')
        );
    }

    public function testAnIntColumnCastsLikePhp(): void
    {
        $path = $this->sheetFile([
            ['id', 'n'],
            ['1', 'abc'],
            ['2', '12abc'],
            ['3', '3.9'],
            ['4', '-2'],
        ]);

        $importer = new class extends Importer {
            protected function getColumns(): array
            {
                return [
                    (object)['field' => 'id'],
                    (object)['field' => 'n', 'type' => 'int'],
                ];
            }
        };

        $this->assertSame([0, 12, 3, -2], array_column($importer->parse($path), 'n'));
    }
}
