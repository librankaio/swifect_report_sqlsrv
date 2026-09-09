<?php

namespace App\Http\Controllers;

use Carbon\Carbon;
use Illuminate\Http\Request;
use App\Exports\PengeluaranExport;
use Illuminate\Support\Facades\DB;
use Maatwebsite\Excel\Facades\Excel;

class PengeluaranController extends Controller
{
    //
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

                    // $results = DB::table('pengeluaran_dokumen')->whereBetween('dptanggal',[$datefrForm,$datetoForm])->where('tstatus','=',1)->where('jenis_dokumen','=',$jenisdok)->get();
                    $query = DB::table('vwLapPengeluaranPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->orderBy('dptanggal','desc')->orderBy('dpnomor','desc');
                    $results = $this->applySearch($query, $request->searchtext)->get();

                    return view('reports.pengeluaran', [
                        'results' => $results
                    ]);
                } else if ($request->jenisdok == "All") {
                    $dtfr = $request->input('dtfrom');
                    $dtto = $request->input('dtto');
                    $jenisdok = $request->input('jenisdok');
                    $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
                    $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

                    // $results = DB::table('vwLapPengeluaranPerDokumenONLINE')->whereBetween('dptanggal',[$datefrForm,$datetoForm])->where('tstatus','=',1)->paginate(10);
                    $query = DB::table('vwLapPengeluaranPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->orderBy('dptanggal','desc')->orderBy('dpnomor','desc');
                    $results = $this->applySearch($query, $request->searchtext)->get();

                    // dd($results);
                    return view('reports.pengeluaran', [
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

                    // $results = DB::table('vwLapPengeluaranPerDokumenONLINE')->whereBetween('dptanggal',[$datefrForm,$datetoForm])->where('tstatus','=',1)->where('jenis_dokumen','=',$jenisdok)->where('dpnomor','=',$searchtext)->paginate(10);
                    $query = DB::table('vwLapPengeluaranPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->orderBy('dptanggal','desc')->orderBy('dpnomor','desc');
                    $results = $this->applySearch($query, $searchtext)->get();

                    return view('reports.pengeluaran', [
                        'results' => $results
                    ]);
                } else if ($request->jenisdok == "All") {
                    $searchtext = $request->searchtext;
                    $dtfr = $request->input('dtfrom');
                    $dtto = $request->input('dtto');
                    $jenisdok = $request->input('jenisdok');
                    $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
                    $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

                    // $results = DB::table('vwLapPengeluaranPerDokumenONLINE')->whereBetween('dptanggal',[$datefrForm,$datetoForm])->where('tstatus','=',1)->where('dpnomor','=',$searchtext)->paginate(10);
                    $query = DB::table('vwLapPengeluaranPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->orderBy('dptanggal','desc')->orderBy('dpnomor','desc');
                    $results = $this->applySearch($query, $searchtext)->get();

                    return view('reports.pengeluaran', [
                        'results' => $results
                    ]);
                }
            }
        }
        return view('reports.pengeluaran');
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

    // public function getPengeluaran(){
    // $data_pengeluaran = Pengeluaran::paginate(1);
    // // return $data_pengeluaran;
    // return view('pengeluaran', [
    //     'data_pengeluaran' => $data_pengeluaran]);
    // }

    // public function showReport(Request $request){
    //     $pengeluaranbtwn = [];
    //     return view('pengeluaran',compact('pengeluaranbtwn'));
    // }

    public function searchPengeluaran(Request $request)
    {

        if ($request->searchtext == null) {
            if ($request->jenisdok != "All") {
                $dtfr = $request->input('dtfrom');
                $dtto = $request->input('dtto');
                $jenisdok = $request->input('jenisdok');
                $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
                $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

                $results = DB::table('vwLapPengeluaranPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->paginate(10);

                return view('pengeluaran', [
                    'results' => $results
                ]);
            } else if ($request->jenisdok == "All") {
                $dtfr = $request->input('dtfrom');
                $dtto = $request->input('dtto');
                $jenisdok = $request->input('jenisdok');
                $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
                $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

                $results = DB::table('vwLapPengeluaranPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->paginate(10);

                return view('reports.pengeluaran', [
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

                $results = DB::table('vwLapPengeluaranPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->paginate(10);

                return view('reports.pengeluaran', [
                    'results' => $results
                ]);
            } else if ($request->jenisdok == "All") {
                $searchtext = $request->searchtext;
                $dtfr = $request->input('dtfrom');
                $dtto = $request->input('dtto');
                $jenisdok = $request->input('jenisdok');
                $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
                $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

                $results = DB::table('vwLapPengeluaranPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->paginate(10);

                return view('reports.pengeluaran', [
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

            $query = DB::table('vwLapPengeluaranPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->orderBy('dptanggal','desc')->orderBy('dpnomor','desc');
            $results = $this->applySearch($query, $request->searchtext)->get();

            // $results = DB::select('EXEC rptTest ?,?,?',[$datefrForm,$datetoForm,$jenisdok]);

            // dd($results);
        } else if ($request->jenisdok == "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');            
            $comp_name = session()->get('comp_name');

            $query = DB::table('vwLapPengeluaranPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->orderBy('dptanggal','desc')->orderBy('dpnomor','desc');
            $results = $this->applySearch($query, $request->searchtext)->get();

            // dd($results);
        }
        return view('print.excel.pengeluaran_report', compact('results', 'datefrForm', 'datetoForm', 'comp_name'));
    }
    public function exportExcelFull(Request $request)
    {
        if ($request->jenisdok != "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');
            $comp_name = session()->get('comp_name');

            $query = DB::table('vwLapPengeluaranPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->orderBy('dptanggal','desc')->orderBy('dpnomor','desc');
            $results = $this->applySearch($query, $request->searchtext)->get();

            // $results = DB::select('EXEC rptTest ?,?,?',[$datefrForm,$datetoForm,$jenisdok]);

            // dd($results);
        } else if ($request->jenisdok == "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');            
            $comp_name = session()->get('comp_name');

            $query = DB::table('vwLapPengeluaranPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->orderBy('dptanggal','desc')->orderBy('dpnomor','desc');
            $results = $this->applySearch($query, $request->searchtext)->get();

            // dd($results);
        }
        return view('print.excel.pengeluaran_report_full', compact('results', 'datefrForm', 'datetoForm', 'comp_name'));
    }

    public function exportPdf(Request $request){
        if ($request->jenisdok != "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

            $query = DB::table('vwLapPengeluaranPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok);
            $results = $this->applySearch($query, $request->searchtext)->get();

            // $results = DB::select('EXEC rptTest ?,?,?', [$datefrForm, $datetoForm, $jenisdok]);

            // dd($results);
        } else if ($request->jenisdok == "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

            $query = DB::table('vwLapPengeluaranPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm]);
            $results = $this->applySearch($query, $request->searchtext)->get();
        }
        return view('print.pdf.pengeluaran_report', compact('results', 'datefrForm', 'datetoForm'));
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

            $query = DB::table('vwLapPengeluaranPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->orderBy('dptanggal','desc')->orderBy('dpnomor','desc');
            $results = $this->applySearch($query, $request->searchtext)->get();
        } else if ($request->jenisdok == "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');
            $comp_name = session()->get('comp_name');

            $query = DB::table('vwLapPengeluaranPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->orderBy('dptanggal','desc')->orderBy('dpnomor','desc');
            $results = $this->applySearch($query, $request->searchtext)->get();
        }

        return Excel::download(new PengeluaranExport($results, $datefrForm, $datetoForm, $comp_name), 'Laporan_PengeluaranDokumen.xlsx');
    }
}
