<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\Validator;
use Illuminate\Support\Str;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class MainController extends Controller
{
    /**
     * Create a new controller instance.
     *
     * @return void
     */
    public function __construct()
    {
//        middleware ...
    }

    public function index()
    {
        return view('index');
    }

    public function loadFile(Request $request)
    {
        $validation = Validator::make($request->all(), [
            'excelSheet' => 'required|file|mimetypes:application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        ]);

        if ($validation->fails()) {
            return view('index')->with(['success' => false, 'errors' => $validation->errors()->all()]);
        }

        $sheet = $request['excelSheet'];

        $reader = IOFactory::createReader('Xlsx');
        $reader->setReadDataOnly(TRUE);
        $spreadsheet = $reader->load($sheet);

        $worksheet = $spreadsheet->getSheet(9);

        $tableaux = [
            ['title_row' => 28, 'start' => 31, 'end' => 45, 'source' => 46],
            ['title_row' => 49, 'start' => 52, 'end' => 149, 'source' => 151],
            ['title_row' => 153, 'start' => 154, 'end' => 159, 'source' => 160],
            ['title_row' => 163, 'start' => 166, 'end' => 192, 'source' => 193],
            ['title_row' => 194, 'start' => 195, 'end' => 199, 'source' => 202],
            ['title_row' => 205, 'start' => 208, 'end' => 212, 'source' => 215],
            ['title_row' => 218, 'start' => 221, 'end' => 229, 'source' => 232],
            ['title_row' => 234, 'start' => 235, 'end' => 248, 'source' => 251],
            ['title_row' => 254, 'start' => 257, 'end' => 264, 'source' => 267],
            ['title_row' => 270, 'start' => 273, 'end' => 284, 'source' => 287],
            ['title_row' => 290, 'start' => 293, 'end' => 300, 'source' => 303],
            ['title_row' => 306, 'start' => 309, 'end' => 312, 'source' => 313],
            ['title_row' => 314, 'start' => 315, 'end' => 317, 'source' => 318],
            ['title_row' => 320, 'start' => 323, 'end' => 326, 'source' => 329],
            ['title_row' => 332, 'start' => 335, 'end' => 345, 'source' => 348],
            ['title_row' => 351, 'start' => 354, 'end' => 357, 'source' => 360],
            ['title_row' => 363, 'start' => 366, 'end' => 396, 'source' => 399],
            ['title_row' => 401, 'start' => 404, 'end' => 408, 'source' => 409],
            ['title_row' => 412, 'start' => 415, 'end' => 417, 'source' => 418],
            ['title_row' => 421, 'start' => 424, 'end' => 431, 'source' => 432],
            ['title_row' => 435, 'start' => 438, 'end' => 445, 'source' => 448],
            ['title_row' => 451, 'start' => 454, 'end' => 457, 'source' => 460],
        ];

//        $finalFileName = "Chapter 11. EAU, ENERGIE ET MINES.xlsx";
        $finalFileName = "Chapter 9. PRODUCTION.xlsx";

        $allData = [];
        foreach ($tableaux as $tableau) {
            $table = [
                'first_row' => [],
                'final' => []
            ];

            foreach ($worksheet->getRowIterator() as $row) {

//                for getting the title
                if ($row->getRowIndex() == $tableau['title_row']) {
                    $cellIterator = $row->getCellIterator();
                    $cellIterator->setIterateOnlyExistingCells(FALSE);
                    $firstTime = true;
                    foreach ($cellIterator as $cell) {
                        if ($firstTime) {
                            array_push($table['first_row'], Str::substr($cell->getValue(), 8, 6));
                            array_push($table['first_row'], Str::substr($cell->getValue(), 14));
//                            $table['first_row'] += ['title' => Str::substr($cell->getValue(), 15)];
//                            $table['first_row'] += ['id_tableau' => Str::substr($cell->getValue(), 8, 7)];
                            $firstTime = false;
                        }
                    }
                }
            }

            foreach ($worksheet->getRowIterator() as $row) {
//                for getting the source
                if ($row->getRowIndex() == $tableau['source']) {
                    $cellIterator = $row->getCellIterator();
                    $cellIterator->setIterateOnlyExistingCells(FALSE);
                    $firstTime = true;
                    foreach ($cellIterator as $cell) {
                        if ($firstTime) {
                            array_push($table['first_row'], Str::substr($cell->getValue(), 8));
//                            $table['first_row'] += ['source' => Str::substr($cell->getValue(), 8)];
                            $firstTime = false;
                        }
                    }
                }
            }

            foreach ($worksheet->getRowIterator() as $row) {
//                for column title
                if ($row->getRowIndex() == $tableau['start']) {
                    $cellIterator = $row->getCellIterator();
                    $cellIterator->setIterateOnlyExistingCells(FALSE);
                    foreach ($cellIterator as $cell) {
                        array_push($table['first_row'], $cell->getValue());
//                        $table['first_row'] += [$cell->getValue() => null];

                    }
                }
            }
//
//            foreach ($worksheet->getRowIterator() as $row) {
////                for table content
//                if ($row->getRowIndex() >= ($tableau['start'] + 1) && $row->getRowIndex() <= $tableau['end']) {
//                    $cellIterator = $row->getCellIterator();
//                    $cellIterator->setIterateOnlyExistingCells(FALSE);
//                    foreach ($cellIterator as $cell) {
//
//                        array_push($table['table_content'], $cell->getValue());
//                    }
//
//                }
//
//            }


            $finalOne = [];
            $dats = [];
            foreach ($worksheet->getRowIterator() as $row) {

                if ($row->getRowIndex() >= ($tableau['start'] + 1) && $row->getRowIndex() <= ($tableau['end'] + 1)) {

                    array_push($finalOne, $table['first_row'][0]);
                    array_push($finalOne, $table['first_row'][1]);
                    array_push($finalOne, $table['first_row'][2]);
                    $cellIterator = $row->getCellIterator();
                    $cellIterator->setIterateOnlyExistingCells(FALSE);
                    foreach ($cellIterator as $cell) {
                        array_push($finalOne, $cell->getValue());
                    }
                    array_push($dats, $finalOne);
                    $finalOne = [];
                }
            }



            $table['final'] = $dats;

            array_push($allData, $table);
        }



//        return response()->json($allData, 200);

        $newSpreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();

        $activeSheet = $newSpreadsheet->getActiveSheet();


        $index = 1;
        $tableItiner = 1;
        foreach ($allData as $tableau) {

//            dump($tableau['first_row']);
            for ($x = 0; $x < count($tableau['first_row']); $x++) {
//                dump($x);
                switch ($x) {
                    case 0:
                        $activeSheet->setCellValueByColumnAndRow($x+1, $index, 'id_tableau');
                        break;
                    case 1:
                        $activeSheet->setCellValueByColumnAndRow($x+1, $index, 'titre_tableau');
                        break;
                    case 2:
                        $activeSheet->setCellValueByColumnAndRow($x+1, $index, 'source');
                        break;
                    default:
                        $activeSheet->setCellValueByColumnAndRow($x+1, $index, $tableau['first_row'][$x]);
                }
            }

            $index++;
            foreach ($tableau['final'] as $final) {

                for ($f = 0; $f < count($final); $f++) {
                    $activeSheet->setCellValueByColumnAndRow($f+1, $index, $final[$f]);
                }

                $index++;
            }
            if (count($allData) != $tableItiner)
            {
                $activeSheet->insertNewRowBefore($index-1, 1);
            }
            $tableItiner++;
        }

        $fileHere = 'mine.xlsx';
        $writer = new Xlsx($newSpreadsheet);
        $writer->save($fileHere);

        return response()->download($fileHere, $finalFileName);
    }
}
