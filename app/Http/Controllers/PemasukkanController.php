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

                    $results = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->orderBy('dptanggal','desc')->orderBy('dpnomor','desc')->get();

                    return view('reports.pemasukkan', [
                        'results' => $results
                    ]);
                } else if ($request->jenisdok == "All") {
                    $dtfr = $request->input('dtfrom');
                    $dtto = $request->input('dtto');
                    $jenisdok = $request->input('jenisdok');
                    $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
                    $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

                    $results = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->orderBy('dptanggal','desc')->orderBy('dpnomor','desc')->get();
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

                    $results = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->orderBy('dptanggal','desc')->orderBy('dpnomor','desc')->where('dpnomor', '=', $searchtext)->get();

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

                    $results = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('dpnomor', '=', $searchtext)->orderBy('dptanggal','desc')->orderBy('dpnomor','desc')->get();

                    return view('reports.pemasukkan', [
                        'results' => $results
                    ]);
                }
            }
        }
        return view('reports.pemasukkan');
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

            $results = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->orderBy('dptanggal','desc')->orderBy('dpnomor','desc')->get();

        } else if ($request->jenisdok == "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');
            $comp_name = session()->get('comp_name');

            $results = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->orderBy('dptanggal','desc')->orderBy('dpnomor','desc')->get();
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

            $results = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->orderBy('dptanggal','desc')->orderBy('dpnomor','desc')->get();

        } else if ($request->jenisdok == "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');
            $comp_name = session()->get('comp_name');

            $results = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->orderBy('dptanggal','desc')->orderBy('dpnomor','desc')->get();
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

            $results = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->get();
        } else if ($request->jenisdok == "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

            $results = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->get();
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

            $results = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->orderBy('dpnomor','asc')->orderBy('dptanggal','asc')->orderBy('bpbnomor','asc')->get();
        } else if ($request->jenisdok == "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');
            $comp_name = session()->get('comp_name');

            $results = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->orderBy('dpnomor','asc')->orderBy('dptanggal','asc')->orderBy('bpbnomor','asc')->get();
        }

        return Excel::download(new PemasukkanExport($results, $datefrForm, $datetoForm, $comp_name), 'Laporan_PemasukanDokumen.xlsx');
    }
}
