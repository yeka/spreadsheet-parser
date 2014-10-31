<?php

namespace Akeneo\Component\SpreadsheetParser\Xls;

class XlsParser
{
    /**
     * @staticvar string the name of the format
     */
    const FORMAT_NAME = 'xls';

    /**
     * @staticvar string Spreadsheet class
     */
    const WORKBOOK_CLASS = 'Akeneo\Component\SpreadsheetParser\Xls\Spreadsheet';

    /**
     * @staticvar string RowIterator class
     */
    const ROW_ITERATOR_CLASS = 'Akeneo\Component\SpreadsheetParser\Xls\RowIterator';

    /**
     * @staticvar string The name of the sheet
     */
    const SHEET_NAME = 'default';

    /**
     * @var SpreadsheetLoader
     */
    private static $spreadsheetLoader;

    /**
     * Opens a CSV file
     *
     * @param string $path
     *
     * @return Spreadsheet
     */
    public static function open($path)
    {
        return static::getSpreadsheetLoader()->open($path);
    }

    /**
     * @return SpreadsheetLoader
     */
    public static function getSpreadsheetLoader()
    {
        if (!isset(self::$spreadsheetLoader)) {
            self::$spreadsheetLoader = new SpreadsheetLoader(
                static::createRowIteratorFactory(),
                static::WORKBOOK_CLASS,
                static::SHEET_NAME
            );
        }

        return self::$spreadsheetLoader;
    }

    /**
     * @return RowIteratorFactory
     */
    protected static function createRowIteratorFactory()
    {
        return new RowIteratorFactory(static::ROW_ITERATOR_CLASS);
    }
}
