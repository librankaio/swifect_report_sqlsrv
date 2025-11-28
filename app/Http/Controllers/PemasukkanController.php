<?php

namespace App\Http\Controllers;

use App\Models\Pemasukkan;
use Barryvdh\DomPDF\PDF as DomPDFPDF;
use Carbon\Carbon;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PDF;

class PemasukkanController extends Controller
{
    public function index(Request $request)
    {
        // dd($request->all());
        if (isset($request->jenisdok)) {
            if ($request->searchtext == null) {
                if ($request->jenisdok != "All") {
                    $dtfr = $request->input('dtfrom');
                    $dtto = $request->input('dtto');
                    $jenisdok = $request->input('jenisdok');
                    $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
                    $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');


                    // $page = request('page', 1);
                    // $pageSize = 10;
                    // $query = DB::select('EXEC rptTest ?,?,?', [$datefrForm, $datetoForm, $jenisdok]);
                    // $offset = ($page * $pageSize) - $pageSize;
                    // $data = array_slice($query, $offset, $pageSize, true);
                    // $results = new \Illuminate\Pagination\LengthAwarePaginator($data, count($data), $pageSize, $page);

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

                    // $results = DB::table('pemasukan_dokumen')->whereBetween('dptanggal',[$datefrForm,$datetoForm])->where('tstatus','=',1)->paginate(10);
                    // $results = DB::table('pemasukan_dokumen')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('tstatus', '=', 1)->get();

                    $results = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->orderBy('dptanggal','desc')->orderBy('dpnomor','desc')->get();
                    // dd($results);
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

                    // $results = DB::table('pemasukan_dokumen')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('tstatus', '=', 1)->where('jenis_dokumen', '=', $jenisdok)->where('dpnomor', '=', $searchtext)->paginate(10);
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

                    // $results = DB::table('pemasukan_dokumen')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('tstatus', '=', 1)->where('dpnomor', '=', $searchtext)->paginate(10);
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

                // $results = DB::table('pemasukan_dokumen')->whereBetween('dptanggal',[$datefrForm,$datetoForm])->where('tstatus','=',1)->wher('jenis_dokumen','=',$jenisdok)->paginate(10);

                // $results = DB::select('EXEC rptTest ?,?,?',[$datefrForm,$datetoForm,$jenisdok]);

                $page = request('page', 1);
                $pageSize = 10;
                $query = DB::select('EXEC rptTest ?,?,?', [$datefrForm, $datetoForm, $jenisdok]);
                $offset = ($page * $pageSize) - $pageSize;
                $data = array_slice($query, $offset, $pageSize, true);
                $results = new \Illuminate\Pagination\LengthAwarePaginator($data, count($data), $pageSize, $page);

                // dd($results);

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
        // dd(request()->all());
        if ($request->jenisdok != "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');
            $comp_name = session()->get('comp_name');

            // $results = DB::table('pemasukan_dokumen')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('stat', '=', 1)->where('jenis_dokumen', '=', $jenisdok)->orderBy('dpnomor','desc')->get();

            $results = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->orderBy('dptanggal','desc')->orderBy('dpnomor','desc')->get();

            // $results = DB::select('EXEC rptTest ?,?,?', [$datefrForm, $datetoForm, $jenisdok]);
            
            // dd("ATAS : ",$results);
        } else if ($request->jenisdok == "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');
            $comp_name = session()->get('comp_name');

            // $results = DB::table('pemasukan_dokumen')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('stat', '=', 1)->orderBy('dptanggal','desc')->get();

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

            // $results = DB::table('pemasukan_dokumen')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('stat', '=', 1)->where('jenis_dokumen', '=', $jenisdok)->orderBy('dpnomor','desc')->get();

            $results = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('jenis_dokumen', '=', $jenisdok)->orderBy('dptanggal','desc')->orderBy('dpnomor','desc')->get();

            // $results = DB::select('EXEC rptTest ?,?,?', [$datefrForm, $datetoForm, $jenisdok]);

            // dd($results);
        } else if ($request->jenisdok == "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');
            $comp_name = session()->get('comp_name');

            // $results = DB::table('pemasukan_dokumen')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('stat', '=', 1)->orderBy('dptanggal','desc')->get();

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


            // if($request->has('download'))
            // {
            //     $pdf = DomPDFPDF::loadView('print.pdf.pemasukkan_report',compact('results'));
            //         return $pdf->stream('pdfview.pdf');
            // }
            // $results = DB::select('EXEC rptTest ?,?,?', [$datefrForm, $datetoForm, $jenisdok]);

            // dd($results);
        } else if ($request->jenisdok == "All") {
            $dtfr = $request->input('dtfrom');
            $dtto = $request->input('dtto');
            $jenisdok = $request->input('jenisdok');
            $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
            $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

            $results = DB::table('vwLapPemasukanPerDokumenONLINE')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->get();

            // if($request->has('download'))
            // {
            //     $pdf = DomPDFPDF::loadView('print.pdf.pemasukkan_report',compact('results'));
            //         return $pdf->stream('pdfview.pdf');    
            // }
        }
        return view('print.pdf.pemasukkan_report', compact('results', 'datefrForm', 'datetoForm'));
    }
}
