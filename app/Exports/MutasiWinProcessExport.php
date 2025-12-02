<?php

namespace App\Exports;

use Illuminate\Support\Collection;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\WithStyles;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use Maatwebsite\Excel\Concerns\WithColumnWidths;
use Maatwebsite\Excel\Concerns\WithColumnFormatting;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Style\Fill;

class MutasiWinProcessExport implements FromCollection, WithHeadings, ShouldAutoSize, WithStyles, WithColumnWidths, WithColumnFormatting
{
    protected $results;
    protected $datefrForm;
    protected $datetoForm;
    protected $comp_name;

    public function __construct($results, $datefrForm, $datetoForm, $comp_name)
    {
        $this->results = $results;
        $this->datefrForm = $datefrForm;
        $this->datetoForm = $datetoForm;
        $this->comp_name = $comp_name;
    }

    public function collection(): Collection
    {
        $data = collect();

        // Add title rows
        $data->push(['LAPORAN PERTANGGUNG JAWABAN POSISI WIP', '', '', '', '', '']);
        $data->push([$this->comp_name, '', '', '', '', '']);

        if ($this->datefrForm && $this->datetoForm) {
            $datefr = date('d/m/Y', strtotime($this->datefrForm));
            $dateto = date('d/m/Y', strtotime($this->datetoForm));
            $data->push(["PERIODE {$datefr} S.D {$dateto}", '', '', '', '', '']);
        }

        $data->push([]); // Empty row

        // Header row
        $data->push([
            'No', 'Kode Barang', 'Nama Barang', 'Satuan', 'Jumlah', 'Keterangan'
        ]);

        if (count($this->results) > 0) {
            $no = 0;

            foreach ($this->results as $item) {
                $no++;
                $data->push([
                    $no,
                    $item->code_mitem,
                    $item->name_mitem,
                    $item->satuan,
                    $item->stock_akhir == 0 ? '--' : (float)$item->stock_akhir,
                    'Sesuai'
                ]);
            }
        } else {
            $data->push(['NO DATA RESULTS...', '', '', '', '', '']);
        }

        $data->push([]); // Empty row
        $data->push(['~ Swifect Inventory BC ~', '', '', '', '', '']);

        return $data;
    }

    public function headings(): array
    {
        return [];
    }

    public function columnWidths(): array
    {
        return [
            'A' => 5,   // No
            'B' => 15,  // Kode Barang
            'C' => 50,  // Nama Barang
            'D' => 8,   // Satuan
            'E' => 12,  // Jumlah
            'F' => 12,  // Keterangan
        ];
    }

    public function styles(Worksheet $sheet)
    {
        // Merge cells for title
        $sheet->mergeCells('A1:F1');
        $sheet->mergeCells('A2:F2');
        if ($this->datefrForm && $this->datetoForm) {
            $sheet->mergeCells('A3:F3');
        }

        // Style the title
        $sheet->getStyle('A1')->getFont()->setBold(true)->setSize(14);
        $sheet->getStyle('A1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        // Style the company name
        $sheet->getStyle('A2')->getFont()->setBold(true)->setSize(12);
        $sheet->getStyle('A2')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        // Style the period
        $sheet->getStyle('A3')->getFont()->setBold(true);
        $sheet->getStyle('A3')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        // Table header should be on row 4
        $tableHeader = 4;

        // Style table header with background color
        $sheet->getStyle("A{$tableHeader}:F{$tableHeader}")->getFont()->setBold(true);
        $sheet->getStyle("A{$tableHeader}:F{$tableHeader}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle("A{$tableHeader}:F{$tableHeader}")->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
        $sheet->getStyle("A{$tableHeader}:F{$tableHeader}")->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setRGB('4F81BD'); // Blue background
        $sheet->getStyle("A{$tableHeader}:F{$tableHeader}")->getFont()->getColor()->setRGB('FFFFFF'); // White text

        // Style data rows with borders and center alignment
        $highestRow = $sheet->getHighestRow();
        for ($row = $tableHeader + 1; $row <= $highestRow; $row++) {
            $sheet->getStyle("A{$row}:F{$row}")->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
            $sheet->getStyle("A{$row}:F{$row}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            // Enable text wrapping for Nama Barang column (C)
            $sheet->getStyle("C{$row}")->getAlignment()->setWrapText(true);
        }

        // Style footer
        $sheet->getStyle("A{$highestRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $sheet->mergeCells("A{$highestRow}:F{$highestRow}");

        return [];
    }

    public function columnFormats(): array
    {
        return [
            'E' => NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1, // Jumlah
        ];
    }
}
