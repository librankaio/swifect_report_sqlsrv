<?php

namespace App\Http\Controllers;

use App\Exports\PemasukkanExport;
use App\Models\Pemasukkan;
use Barryvdh\DomPDF\PDF as DomPDFPDF;
use Carbon\Carbon;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use Maatwebsite\Excel\Facades\Excel;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PDF;

class PemasukkanController extends Controller
{
    public function index(Request $request)
    {
        if (isset($request->jenisdok)) {
            if ($request->searchtext == null) {
                if ($request->jenisdok != "All") {
                    $dtfr = $request->input('dtfrom');
                    $dtto = $request->input('dtto');
                    $jenisdok = $request->input('jenisdok');
                    $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
                    $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

                    $query = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->orderBy('dptanggal','desc')->orderBy('dpnomor','desc');
                    $results = $this->applySearch($query, $request->searchtext)->get();

                    return view('reports.pemasukkan', [
                        'results' => $results
                    ]);
                } else if ($request->jenisdok == "All") {
                    $dtfr = $request->input('dtfrom');
                    $dtto = $request->input('dtto');
                    $jenisdok = $request->input('jenisdok');
                    $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
                    $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

                    $query = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->orderBy('dptanggal','desc')->orderBy('dpnomor','desc');
                    $results = $this->applySearch($query, $request->searchtext)->get();
                    return view('reports.pemasukkan', [
                        'results' => $results
                    ]);
                }
            } else if ($request->searchtext != null) {
                if ($request->jenisdok != "All") {
                    $searchtext = $request->searchtext;
                    $dtfr = $request->input('dtfrom');
                    $dtto = $request->input('dtto');
                    $jenisdok = $request->input('jenisdok');
                    $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
                    $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

                    $query = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->orderBy('dptanggal','desc')->orderBy('dpnomor','desc');
                    $results = $this->applySearch($query, $searchtext)->get();

                    return view('reports.pemasukkan', [
                        'results' => $results
                    ]);
                } else if ($request->jenisdok == "All") {
                    $searchtext = $request->searchtext;
                    $dtfr = $request->input('dtfrom');
                    $dtto = $request->input('dtto');
                    $jenisdok = $request->input('jenisdok');
                    $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
                    $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

                    $query = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->orderBy('dptanggal','desc')->orderBy('dpnomor','desc');
                    $results = $this->applySearch($query, $searchtext)->get();

                    return view('reports.pemasukkan', [
                        'results' => $results
                    ]);
                }
            }
        }
        return view('reports.pemasukkan');
    }

    private function applySearch($query, $searchtext)
    {
        $searchtext = trim((string) $searchtext);

        if ($searchtext === '') {
            return $query;
        }

        $columns = $this->searchableColumns($query->from);

        if (empty($columns)) {
            return $query;
        }

        // Netralkan wildcard LIKE supaya "100%" tidak cocok ke semua baris
        $like = '%' . str_replace(['\\', '%', '_', '['], ['\\\\', '\\%', '\\_', '\\['], $searchtext) . '%';

        $dateTypes = ['date', 'datetime', 'datetime2', 'smalldatetime', 'datetimeoffset'];

        return $query->where(function ($q) use ($columns, $like, $dateTypes) {
            foreach ($columns as $name => $type) {
                $safe = str_replace(']', ']]', $name);

                if (in_array($type, $dateTypes)) {
                    $q->orWhereRaw("CONVERT(VARCHAR(10), [{$safe}], 103) LIKE ? ESCAPE '\\'", [$like]); // dd/mm/yyyy
                    $q->orWhereRaw("CONVERT(VARCHAR(10), [{$safe}], 23) LIKE ? ESCAPE '\\'", [$like]);  // yyyy-mm-dd
                } else {
                    $q->orWhereRaw("CAST([{$safe}] AS NVARCHAR(200)) LIKE ? ESCAPE '\\'", [$like]);
                }
            }
        });
    }

    private function searchableColumns($table)
    {
        static $cache = [];

        if (!array_key_exists($table, $cache)) {
            // Tipe yang tidak masuk akal / tidak aman untuk LIKE
            $skipped = ['binary', 'varbinary', 'image', 'xml', 'geography', 'geometry', 'hierarchyid', 'timestamp', 'sql_variant'];

            $columns = [];

            foreach (DB::select('SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = ?', [$table]) as $column) {
                $type = strtolower($column->DATA_TYPE);

                if (!in_array($type, $skipped)) {
                    $columns[$column->COLUMN_NAME] = $type;
                }
            }

            $cache[$table] = $columns;
        }

        return $cache[$table];
    }

    public function searchPemasukan(Request $request)
    {
        if ($request->searchtext == null) {
            if ($request->jenisdok != "All") {
                $dtfr = $request->input('dtfrom');
                $dtto = $request->input('dtto');
                $jenisdok = $request->input('jenisdok');
                $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
                $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

                $page = request('page', 1);
                $pageSize = 10;
                $query = DB::select('EXEC rptTest ?,?,?', [$datefrForm, $datetoForm, $jenisdok]);
                $offset = ($page * $pageSize) - $pageSize;
                $data = array_slice($query, $offset, $pageSize, true);
                $results = new \Illuminate\Pagination\LengthAwarePaginator($data, count($data), $pageSize, $page);

                return view('reports.pemasukkan', [
                    'results' => $results
                ]);
            } else if ($request->jenisdok == "All") {
                $dtfr = $request->input('dtfrom');
                $dtto = $request->input('dtto');
                $jenisdok = $request->input('jenisdok');
                $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
                $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

                $results = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('tstatus', '=', 1)->paginate(10);

                return view('reports.pemasukkan', [
                    'results' => $results
                ]);
            }
        } else if ($request->searchtext != null) {
            if ($request->jenisdok != "All") {
                $searchtext = $request->searchtext;
                $dtfr = $request->input('dtfrom');
                $dtto = $request->input('dtto');
                $jenisdok = $request->input('jenisdok');
                $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
                $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

                $results = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('tstatus', '=', 1)->where('jenis_dokumen', '=', $jenisdok)->where('dpnomor', '=', $searchtext)->paginate(10);

                return view('reports.pemasukkan', [
                    'results' => $results
                ]);
            } else if ($request->jenisdok == "All") {
                $searchtext = $request->searchtext;
                $dtfr = $request->input('dtfrom');
                $dtto = $request->input('dtto');
                $jenisdok = $request->input('jenisdok');
                $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
                $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

                $results = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('tstatus', '=', 1)->where('dpnomor', '=', $searchtext)->paginate(10);

                return view('reports.pemasukkan', [
                    'results' => $results
                ]);
            }
        }
    }

    public function exportExcel(Request $request)
    {
        if ($request->jenisdok != "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');
            $comp_name = session()->get('comp_name');

            $query = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->orderBy('dptanggal','desc')->orderBy('dpnomor','desc');
            $results = $this->applySearch($query, $request->searchtext)->get();

        } else if ($request->jenisdok == "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');
            $comp_name = session()->get('comp_name');

            $query = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->orderBy('dptanggal','desc')->orderBy('dpnomor','desc');
            $results = $this->applySearch($query, $request->searchtext)->get();
        }
        return view('print.excel.pemasukkan_report', compact('results', 'datefrForm', 'datetoForm', 'comp_name'));
    }
    public function exportExcelFull(Request $request)
    {
        // dd(request()->all());
        if ($request->jenisdok != "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');
            $comp_name = session()->get('comp_name');

            $query = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->orderBy('dptanggal','desc')->orderBy('dpnomor','desc');
            $results = $this->applySearch($query, $request->searchtext)->get();

        } else if ($request->jenisdok == "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');
            $comp_name = session()->get('comp_name');

            $query = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->orderBy('dptanggal','desc')->orderBy('dpnomor','desc');
            $results = $this->applySearch($query, $request->searchtext)->get();
        }
        return view('print.excel.pemasukkan_report_full', compact('results', 'datefrForm', 'datetoForm', 'comp_name'));
    }

    public function exportPdf(Request $request){
        if ($request->jenisdok != "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

            $query = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok);
            $results = $this->applySearch($query, $request->searchtext)->get();
        } else if ($request->jenisdok == "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

            $query = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm]);
            $results = $this->applySearch($query, $request->searchtext)->get();
        }
        return view('print.pdf.pemasukkan_report', compact('results', 'datefrForm', 'datetoForm'));
    }

    public function exportExcel2(Request $request)
    {
        if ($request->jenisdok != "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');
            $comp_name = session()->get('comp_name');

            $query = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->orderBy('dpnomor','asc')->orderBy('dptanggal','asc')->orderBy('bpbnomor','asc');
            $results = $this->applySearch($query, $request->searchtext)->get();
        } else if ($request->jenisdok == "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');
            $comp_name = session()->get('comp_name');

            $query = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->orderBy('dpnomor','asc')->orderBy('dptanggal','asc')->orderBy('bpbnomor','asc');
            $results = $this->applySearch($query, $request->searchtext)->get();
        }

        return Excel::download(new PemasukkanExport($results, $datefrForm, $datetoForm, $comp_name), 'Laporan_PemasukanDokumen.xlsx');
    }
}
