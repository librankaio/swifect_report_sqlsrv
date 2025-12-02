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

class PemasukkanExport implements FromCollection, WithHeadings, ShouldAutoSize, WithStyles, WithColumnWidths, WithColumnFormatting
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
        $data->push(['LAPORAN PERTANGGUNG JAWABAN PEMASUKAN DOKUMEN', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        $data->push([$this->comp_name, '', '', '', '', '', '', '', '', '', '', '', '', '', '']);

        if ($this->datefrForm && $this->datetoForm) {
            $datefr = date('d/m/Y', strtotime($this->datefrForm));
            $dateto = date('d/m/Y', strtotime($this->datetoForm));
            $data->push(["PERIODE {$datefr} S.D {$dateto}", '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        }

        $data->push([]); // Empty row

        // Header row 1 (with colspans)
        $data->push([
            'No', 'Jenis Dokumen', 'Nomor Aju', 'Dokumen Pabean', '', 'Bukti Penerimaan Barang', '', 'Supplier',
            'Kode Barang', 'Nama Barang', 'Satuan', 'Jumlah', 'Nilai Barang', ''
        ]);

        // Header row 2 (sub-headers)
        $data->push([
            '', '', '', 'Nomor Pendaftaran', 'Tanggal', 'Nomor', 'Tanggal', '',
            '', '', '', '', 'Rupiah', 'USD'
        ]);

        if ($this->results->count() > 0) {
            $no = 0;
            $dpnomor = '';
            $bpbnomor = '';

            foreach ($this->results as $item) {
                if ($item->dpnomor == $dpnomor) {
                    // For merged rows, empty first 7 columns
                    $data->push([
                        '', '', '', '', '', '', '', '',
                        $item->kode_barang,
                        $item->nama_barang,
                        $item->sat,
                        $item->jumlah == 0 ? '0' : (float)$item->jumlah,
                        $item->nilai_barang == 0 ? '0' : (float)$item->nilai_barang,
                        $item->nilai_barang_usd == 0 ? '0' : (float)$item->nilai_barang_usd
                    ]);
                } elseif ($item->dpnomor == $dpnomor && $item->bpbnomor != $bpbnomor) {
                    // Special case for different bpbnomor but same dpnomor
                    $data->push([
                        '', '', '', '', '', $item->bpbnomor, date('d/m/Y', strtotime($item->bpbtanggal)), $item->pembeli_penerima,
                        $item->kode_barang,
                        $item->nama_barang,
                        $item->sat,
                        $item->jumlah == 0 ? '0' : (float)$item->jumlah,
                        $item->nilai_barang == 0 ? '0' : (float)$item->nilai_barang,
                        $item->nilai_barang_usd == 0 ? '0' : (float)$item->nilai_barang_usd
                    ]);
                } else {
                    // New document number
                    $no++;
                    $dpnomor = $item->dpnomor;
                    $bpbnomor = $item->bpbnomor;

                    $data->push([
                        $no,
                        $item->jenis_dokumen,
                        $item->nomoraju,
                        $item->dpnomor,
                        date('d/m/Y', strtotime($item->dptanggal)),
                        $item->bpbnomor,
                        date('d/m/Y', strtotime($item->bpbtanggal)),
                        $item->pemasok_pengirim,
                        $item->kode_barang,
                        $item->nama_barang,
                        $item->sat,
                        $item->jumlah == 0 ? '0' : (float)$item->jumlah,
                        $item->nilai_barang == 0 ? '0' : (float)$item->nilai_barang,
                        $item->nilai_barang_usd == 0 ? '0' : (float)$item->nilai_barang_usd
                    ]);
                }
            }
        } else {
            $data->push(['NO DATA RESULTS...', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        }

        $data->push([]); // Empty row
        $data->push(['~ Swifect Inventory BC ~', '', '', '', '', '', '', '', '', '', '', '', '', '']);

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
            'B' => 15,  // Jenis Dokumen
            'C' => 15,  // Nomor Aju
            'D' => 20,  // Nomor Pendaftaran
            'E' => 12,  // Tanggal Dokumen
            'F' => 20,  // Nomor Bukti
            'G' => 12,  // Tanggal Bukti
            'H' => 25,  // Supplier
            'I' => 15,  // Kode Barang
            'J' => 50,  // Nama Barang
            'K' => 8,   // Satuan
            'L' => 12,  // Jumlah
            'M' => 18,  // Nilai Rp
            'N' => 18,  // Nilai USD
        ];
    }

    public function styles(Worksheet $sheet)
    {
        // Merge cells for title
        $sheet->mergeCells('A1:N1');
        $sheet->mergeCells('A2:N2');
        if ($this->datefrForm && $this->datetoForm) {
            $sheet->mergeCells('A3:N3');
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

        // Table headers should be on rows 4 and 5
        $tableHeader1 = 4; // First header row
        $tableHeader2 = 5; // Second header row

        // Merge header cells to match colspan
        $sheet->mergeCells("D{$tableHeader1}:E{$tableHeader1}"); // Dokumen Pabean
        $sheet->mergeCells("F{$tableHeader1}:G{$tableHeader1}"); // Bukti Penerimaan Barang
        $sheet->mergeCells("M{$tableHeader1}:N{$tableHeader1}"); // Nilai Barang

        // Style table headers with background color
        $sheet->getStyle("A{$tableHeader1}:N{$tableHeader2}")->getFont()->setBold(true);
        $sheet->getStyle("A{$tableHeader1}:N{$tableHeader2}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle("A{$tableHeader1}:N{$tableHeader2}")->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
        $sheet->getStyle("A{$tableHeader1}:N{$tableHeader2}")->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setRGB('4F81BD'); // Blue background
        $sheet->getStyle("A{$tableHeader1}:N{$tableHeader2}")->getFont()->getColor()->setRGB('FFFFFF'); // White text

        // Style data rows with borders and center alignment
        $highestRow = $sheet->getHighestRow();
        for ($row = $tableHeader2 + 1; $row <= $highestRow; $row++) {
            $sheet->getStyle("A{$row}:N{$row}")->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
            $sheet->getStyle("A{$row}:N{$row}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            // Enable text wrapping for Supplier column (H) and Nama Barang column (J)
            $sheet->getStyle("H{$row}")->getAlignment()->setWrapText(true);
            $sheet->getStyle("J{$row}")->getAlignment()->setWrapText(true);
        }

        // Style footer
        $sheet->getStyle("A{$highestRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $sheet->mergeCells("A{$highestRow}:N{$highestRow}");

        return [];
    }

    public function columnFormats(): array
    {
        return [
            'L' => NumberFormat::FORMAT_NUMBER_00, // Jumlah column - 2 decimal places
            'M' => NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1, // Nilai Rp column - with commas
            'N' => NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1, // Nilai USD column - with commas
        ];
    }
}
