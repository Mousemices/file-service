<?php


namespace App\Http\Controllers\API;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Storage;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Exception;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class ExcelController extends Controller
{
    protected int $index = 5; // Ilość Sprzedaż 销售数量

    protected int $value = 10;

    // TODO: Use UpdateExcelRequest instead

    /**
     * @throws Exception
     */
    public function update(Request $request)
    {
        $excelFile = $request->file('excel_file');

        $originalName = $excelFile->getClientOriginalName();

        $spreadsheet = IOFactory::load($excelFile->getPathname());

        $worksheet = $spreadsheet->getActiveSheet();
        $rows = $worksheet->toArray();

        $isFirstRow = true;
        $headers = [];
        $filteredRows = [];
        $ceiled = true;

        foreach ($rows as $row) {
            if ($isFirstRow) {
                $headers[] = $row;
                $isFirstRow = false;
                continue;
            }

            $count = $row[$this->index];
            if ($count > $this->value) {
                if ($ceiled) {
                    $row[$this->index] = $this->integerRoundUp(intval($row[$this->index]));
                }

                $filteredRows[] = $row;
            }
        }

        $spreadsheet = new Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();
        // Add headers to the first row
        $worksheet->fromArray($headers);

        for ($i = 0; $i < count($filteredRows); $i++) {
            $rowData = array_values($filteredRows[$i]); // Convert associative array to indexed array
            $rowIndex = $i + 2; // Excel rows are 1 - indexed
            $worksheet->fromArray([$rowData], null, 'A' . $rowIndex);
        }

        // Create a new Xlsx writer
        $writer = new Xlsx($spreadsheet);

        $uniqueFileName = uniqid() . '-' . $originalName;
        $path = storage_path('app/public/' . $uniqueFileName);

        try {
            $writer->save($path);
        } catch (\Exception) {
            return response()->json(['message' => 'Server Internal Error']);
        }

        $destinationPath = public_path('storage/' . $uniqueFileName);

        return response()->download($destinationPath)->deleteFileAfterSend();
    }

    protected function integerRoundUp(int $number): int
    {
        $remainder = $number % 10;
        if ($remainder === 0) {
            return $number; // No rounding needed, already a multiple of 10
        }

        return $number + (10 - $remainder);
    }
}
