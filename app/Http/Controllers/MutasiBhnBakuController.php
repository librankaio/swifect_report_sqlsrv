<?php

namespace App\Http\Controllers;

use Carbon\Carbon;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;

class MutasiBhnBakuController extends Controller
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

                // $results = DB::select('CALL rptmutasibahanbaku (?,?,?)', [$datefrForm, $datetoForm, $compcode]);
                $results = DB::select('EXEC LapMutasiBahanBakuOCIOnline ?,?', [$datefrForm, $datetoForm]);

                // $query = DB::select('EXEC rptTest ?,?,?',[$datefrForm,$datetoForm,'BC 4.0']);

                $page = request('page', 1);
                $pageSize = 25;
                // $query = DB::select('CALL rptmutasibahanbaku (?,?,?)', [$datefrForm, $datetoForm, $compcode]);
                $query = DB::select('EXEC LapMutasiBahanBakuOCIOnline ?,?', [$datefrForm, $datetoForm]);
                $offset = ($page * $pageSize) - $pageSize;
                $data = array_slice($query, $offset, $pageSize, true);
                // $results = new \Illuminate\Pagination\LengthAwarePaginator($data, count($data), $pageSize, $page);

                // dd($results);

                return view('reports.mutasibhnbaku', [
                    'results' => $results
                ]);
            } else if ($request->searchtext != null) {
                $searchtext = trim($request->searchtext);
                $dtfr = $request->input('dtfrom');
                $dtto = $request->input('dtto');
                $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
                $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');

                $rows = DB::select('EXEC LapMutasiBahanBakuOCIOnline ?,?', [$datefrForm, $datetoForm]);

                $results = array_values(array_filter($rows, function ($row) use ($searchtext) {
                    return stripos($row->code_mitem ?? '', $searchtext) !== false
                        || stripos($row->name_mitem ?? '', $searchtext) !== false;
                }));

                return view('reports.mutasibhnbaku', [
                    'results' => $results
                ]);
            }
        }
        return view('reports.mutasibhnbaku');
    }

    public function exportExcel(Request $request)
    {
        $dtfr = $request->input('dtfrom');
        $dtto = $request->input('dtto');
        $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
        $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');
        $comp_code = session()->get('comp_code');
        $comp_name = session()->get('comp_name');

        $results = DB::select('EXEC LapMutasiBahanBakuOCIOnline ?,?', [$datefrForm, $datetoForm]);

        // dd($results);

        return view('print.excel.mutasibhnbaku_report', compact('results', 'datefrForm', 'datetoForm', 'comp_name'));
    }

    public function exportPdf(Request $request)
    {
        $dtfr = $request->input('dtfrom');
        $dtto = $request->input('dtto');
        $datefrForm = Carbon::createFromFormat('d/m/Y', $dtfr)->format('Y-m-d');
        $datetoForm = Carbon::createFromFormat('d/m/Y', $dtto)->format('Y-m-d');
        $compcode = session()->get('comp_code');

        $results = DB::select('EXEC LapMutasiBahanBakuOCIOnline ?,?', [$datefrForm, $datetoForm]);

        // dd($results);

        return view('print.pdf.mutasibhnbaku_report', compact('results', 'datefrForm', 'datetoForm', 'compcode'));
    }
}
