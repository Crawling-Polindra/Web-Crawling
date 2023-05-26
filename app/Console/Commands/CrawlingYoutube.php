<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use Illuminate\Support\Facades\Http;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as ReaderXlsx;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class CrawlingYoutube extends Command
{
    const FILE = "new-youtube.xlsx";

    protected $signature = 'app:crawling-youtube';

    protected $description = 'Command description';

    public function handle()
    {
        $response = Http::get("https://www.googleapis.com/youtube/v3/search", [
            "part" => "snippet",
            "key" => "AIzaSyDLPtm7PG-GPeuOQmV8ogGYrPDcZBq7AO4",
            "q" => "Politik",
            "maxResults" => 50
        ]);

        $spreadsheet = new Spreadsheet();
        $activeWorksheet = $spreadsheet->getActiveSheet();

        $index = 1;


        if (is_file(public_path(self::FILE))) {
            $this->newRowData($response->json()['items']);
            return;
        }
        $activeWorksheet->setCellValue("A$index", "publishedAt");
        $activeWorksheet->setCellValue("B$index", "channelId");
        $activeWorksheet->setCellValue("C$index", "title");
        $activeWorksheet->setCellValue("D$index", "description");
        $activeWorksheet->setCellValue("E$index", "channelTitle");

        foreach ($response->json()['items'] as $key => $item) {
            $index++;

            $activeWorksheet->setCellValue("A$index", "publishedAt");
            $activeWorksheet->setCellValue("B$index", "channelId");
            $activeWorksheet->setCellValue("C$index", "title");
            $activeWorksheet->setCellValue("D$index", "description");
            $activeWorksheet->setCellValue("E$index", "channelTitle");


            if (is_file(public_path(self::FILE))) {
                $this->newRowData($response->json()['items']);
                return;
            }

            foreach ($response->json()['items'] as $key => $item) {
                $index++;

                $activeWorksheet->setCellValue("A$index", $item['snippet']['publishedAt']);
                $activeWorksheet->setCellValue("B$index", $item['snippet']['channelId']);
                $activeWorksheet->setCellValue("C$index", $item['snippet']['title']);
                $activeWorksheet->setCellValue("D$index", $item['snippet']['description']);
                $activeWorksheet->setCellValue("E$index", $item['snippet']['channelTitle']);
            }

            $writer = new Xlsx($spreadsheet);
            $writer->save(public_path(self::FILE));
            return;
        }
    }

    protected function newRowData($data)
    {
        $render = new ReaderXlsx();
        $spreadsheet = $render->load(public_path(self::FILE));
        $activeWorksheet = $spreadsheet->getActiveSheet();
        $newRow = $activeWorksheet->getHighestRow();

        foreach ($data as $key => $item) {
            $newRow++;

            $activeWorksheet->setCellValue("A$newRow", $item['snippet']['publishedAt']);
            $activeWorksheet->setCellValue("B$newRow", $item['snippet']['channelId']);
            $activeWorksheet->setCellValue("C$newRow", $item['snippet']['title']);
            $activeWorksheet->setCellValue("D$newRow", $item['snippet']['description']);
            $activeWorksheet->setCellValue("E$newRow", $item['snippet']['channelTitle']);
        }

        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save(public_path(self::FILE));
    }
}
