<?php

namespace App\Http\Controllers;

use Carbon\Carbon;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;

class MutasiWinProcessController extends Controller
{
    public function index(Request $request)
    {
        if (isset($request->dtfrom)) {
            if ($request->searchtext == null) {

                $dtfr = $request->input('dtfrom');
                $dtto = $request->input('dtto');
                $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
                $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');
                $compcode = session()->get('comp_code');

                $results = DB::select('CALL rptmutasiwip (?,?,?)', [$datefrForm, $datetoForm, $compcode]);

                // $query = DB::select('EXEC rptTest ?,?,?',[$datefrForm,$datetoForm,'BC 4.0']);

                $page = request('page', 1);
                $pageSize = 25;
                $query = DB::select('CALL rptmutasiwip (?,?,?)', [$datefrForm, $datetoForm, $compcode]);
                $offset = ($page * $pageSize) - $pageSize;
                $data = array_slice($query, $offset, $pageSize, true);
                // $results = new \Illuminate\Pagination\LengthAwarePaginator($data, count($data), $pageSize, $page);

                // dd($results);

                return view('reports.mutasiwinprocess', [
                    'results' => $results
                ]);
            } else if ($request->searchtext != null) {
                $searchtext = $request->searchtext;
                $dtfr = $request->input('dtfrom');
                $dtto = $request->input('dtto');
                $jenisdok = $request->input('jenisdok');
                $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
                $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

                $results = DB::table('pemasukan_dokumen')->whereBetween('dptanggal', [$datefrForm, $datetoForm])->where('tstatus', '=', 1)->where('jenis_dokumen', '=', $jenisdok)->where('dpnomor', '=', $searchtext)->paginate(10);

                return view('reports.mutasiwinprocess', [
                    'results' => $results
                ]);
            }
        }
        return view('reports.mutasiwinprocess');
    }

    public function exportExcel(Request $request)
    {
        $dtfr = $request->input('dtfrom');
        $dtto = $request->input('dtto');
        $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
        $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');
        $comp_code = session()->get('comp_code');
        $comp_name = session()->get('comp_name');

        $results = DB::select('CALL rptmutasiwip (?,?,?)', [$datefrForm, $datetoForm, $comp_code]);

        // dd($results);

        return view('print.excel.mutasiwinprocess_report', compact('results', 'datefrForm', 'datetoForm', 'comp_name'));
    }

    public function exportPdf(Request $request)
    {
        $dtfr = $request->input('dtfrom');
        $dtto = $request->input('dtto');
        $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
        $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');
        $compcode = session()->get('comp_code');
        $comp_name = session()->get('comp_name');

        $results = DB::select('CALL rptmutasiwip (?,?,?)', [$datefrForm, $datetoForm, $compcode]);

        // dd($results);

        return view('print.pdf.mutasiwinprocess_report', compact('results', 'datefrForm', 'datetoForm', 'compcode','comp_name'));
    }
}
