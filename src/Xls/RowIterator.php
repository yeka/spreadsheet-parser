<?php

namespace Akeneo\Component\SpreadsheetParser\Xls;

use DateTime;
use Symfony\Component\OptionsResolver\OptionsResolver;
use Symfony\Component\OptionsResolver\OptionsResolverInterface;

/**
 * Row iterator for CSV
 *
 * The following options are available :
 *  - length:    the maximum length of read lines
 *  - delimiter: the CSV delimiter character
 *  - enclosure: the CSV enclosure character
 *  - escape:    the CSV escape character
 *  - encoding:  the encoding of the CSV file
 *
 * @author    Antoine Guigan <antoine@akeneo.com>
 * @copyright 2014 Akeneo SAS (http://www.akeneo.com)
 * @license   http://opensource.org/licenses/osl-3.0.php  Open Software License (OSL 3.0)
 */
class RowIterator implements \Iterator
{
    /**
     * @var string
     */
    protected $path;

    /**
     * @var array
     */
    protected $options;

    /**
     * @var SpreadsheetExcelReader
     */
    protected $xls;

    /**
     * @var array
     */
    protected $currentKey;

    /**
     * @var array
     */
    protected $currentValue;

    /**
     * @var boolean
     */
    protected $valid;

    protected $sheet_index;

    /**
     * Constructor
     *
     * @param string $path
     * @param array  $options
     */
    public function __construct(
        $path,
        $sheet_index,
        array $options
    ) {
        $this->path = $path;
        $resolver = new OptionsResolver;
        $this->setDefaultOptions($resolver);
        $this->options = $resolver->resolve($options);
        $this->sheet_index = $sheet_index;
    }

    /**
     * {@inheritdoc}
     */
    public function current()
    {
        $rows = array();
        for ($col = 1; $col <= $this->xls->colcount($this->sheet_index); $col++) {
            $type = $this->xls->type($this->currentKey, $col, $this->sheet_index);

            switch ($type) {
                case 'number':
                    $val = $this->xls->raw($this->currentKey, $col, $this->sheet_index);
                    break;
                case 'date':
                    $raw = $this->xls->raw($this->currentKey, $col, $this->sheet_index);
                    // Converting Excel Raw Date
                    $days = floor($raw);
                    $secs = round(($raw - $days) * 24 * 3600);
                    $date = new DateTime('1899-12-30');

                    $date->modify("+{$days} day");
                    $date->modify("+{$secs} second");

                    $val = $date;
                    break;
                default:
                    $val = $this->xls->val($this->currentKey, $col, $this->sheet_index);
            }
            $rows[] = $val;
        }
        return $rows;
    }

    /**
     * {@inheritdoc}
     */
    public function key()
    {
        return $this->currentKey;
    }

    /**
     * {@inheritdoc}
     */
    public function next()
    {
        $this->currentKey++;
        $this->valid = ($this->currentKey <= $this->xls->rowcount($this->sheet_index));
    }

    /**
     * {@inheritdoc}
     */
    public function rewind()
    {
        if ($this->xls) {
            rewind($this->xls);
        } else {
            $this->openResource();
        }
        $this->currentKey = 0;
        $this->next();
    }

    /**
     * {@inheritdoc}
     */
    public function valid()
    {
        return $this->valid;
    }

    /**
     * Sets the default options
     *
     * @param OptionsResolverInterface $resolver
     */
    protected function setDefaultOptions(OptionsResolverInterface $resolver)
    {
        $resolver->setOptional(['encoding']);
        $resolver->setDefaults(
            [
                'length'    => null,
                'delimiter' => ',',
                'enclosure' => '"',
                'escape'    => '\\'
            ]
        );
    }

    /**
     * Opens the file resource
     *
     * @return resource
     */
    protected function openResource()
    {
        $this->xls = new SpreadsheetExcelReader($this->path);
    }

    /**
     * Returns the server encoding
     *
     * @return string
     */
    protected function getCurrentEncoding()
    {
        $locale = explode('.', setlocale(LC_CTYPE, 0));

        return isset($locale[1]) ? $locale[1] : 'UTF8';
    }
}
