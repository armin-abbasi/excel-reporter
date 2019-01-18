<?php

namespace Xerobase\ExcelReporter;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Csv;

class Export
{

    /**
     * @var $sheet
     */
    private $sheet;

    /**
     * @var $spreadSheet
     */
    private $spreadSheet;

    /**
     * @var $format
     */
    private $format = 'xlsx';

    /**
     * @var $allowedFormats
     */
    private $allowedFormats = ['csv', 'xlsx'];

    /**
     * @var $excludes
     */
    private $excludes;

    /**
     * @var $rightToLeft
     */
    private $rightToLeft = false;

    public function __construct()
    {
        $this->spreadSheet = new Spreadsheet();
    }

    /**
     * @param $data
     * @return string
     */
    public function export($data)
    {
        // Convert to array, if object given
        $data = $this->convertToArray($data);

        // Retrieve head cell keys and remove excluded
        $headCells = $this->getHeadCellKeys($data, $this->excludes);

        try {
            // Setting default style
            $this->spreadSheet->getDefaultStyle()->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);

            $this->sheet = $this->spreadSheet->getActiveSheet()->setRightToLeft($this->rightToLeft);
        } catch (\Exception $e) {
            return $e->getMessage();
        }

        // Set header cells
        $colCounter = 0;
        foreach ($headCells as $head) {
            $this->sheet->setCellValue($this->getNameFromNumber($colCounter) . '1', $head);
            $colCounter++;
        }

        // Set body cells
        $rowCounter = 2;
        foreach ($data as $body) {
            $colCounter = 0;
            foreach ($headCells as $head) {
                $this->sheet->setCellValue($this->getNameFromNumber($colCounter) . $rowCounter, $body[$head]);
                $this->sheet->getColumnDimension($this->getNameFromNumber($colCounter))->setAutoSize(true);
                $colCounter++;
            }
            $rowCounter++;
        }

        try {
            // Export the file
            switch ($this->format) {
                case 'csv':
                    $writer = new Csv($this->spreadSheet);
                    header('Content-Disposition: attachment; filename="report.csv"');
                    $writer->save('php://output');
                    break;

                case 'xlsx':
                    $writer = new Xlsx($this->spreadSheet);
                    header('Content-type: application/vnd.ms-excel');
                    header('Content-Disposition: attachment; filename="report.xlsx"');
                    $writer->save('php://output');
                    break;
            }

        } catch (\Exception $e) {
            return $e->getMessage();
        }

        return 'Something goes wrong :(';
    }

    /**
     * @param string $format
     * @return $this
     */
    public function setFormat($format)
    {
        $format = strtolower($format);
        // If user format is allowed take it, other wise return Microsoft Excel format
        $this->format = in_array($format, $this->allowedFormats) ? $format : 'xlsx';

        return $this;
    }

    /**
     * @param $excludes
     * @return $this
     */
    public function filterColumns($excludes)
    {
        $this->excludes = $excludes;

        return $this;
    }

    /**
     * Set worksheet direction to RTL
     * @return $this
     */
    public function setRightToLeft()
    {
        $this->rightToLeft = true;

        return $this;
    }


    /**
     * @param $num
     * @return string
     */
    private function getNameFromNumber($num)
    {
        $numeric = $num % 26;
        $letter = chr(65 + $numeric);
        $num2 = intval($num / 26);

        if ($num2 > 0) {
            return $this->getNameFromNumber($num2 - 1) . $letter;
        } else {
            return $letter;
        }
    }

    /**
     * @param $input
     * @return mixed
     */
    private function convertToArray($input)
    {
        // If object given, convert it to array
        if (is_object($input)) {
            $input = json_decode(json_encode($input), true);
        }

        // Make it two dimensional
        if (empty($input[0]) || !is_array($input[0])) {
            $input = [
                $input
            ];
        }

        return $input;
    }

    /**
     * @param $input
     * @param $excludes
     * @return array
     */
    private function getHeadCellKeys($input, $excludes)
    {
        $keys = array_keys($input[0]);

        if (!empty($excludes)) {
            $excludes = (array)$excludes;

            $keys = array_filter($keys, function ($key) use ($excludes) {
                if (!in_array($key, $excludes)) {
                    return $key;
                }

                return null;
            });
        }

        return $keys;
    }

}