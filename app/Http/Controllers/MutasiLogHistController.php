<?php

namespace App\Http\Controllers;

use Carbon\Carbon;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;

class MutasiLogHistController extends Controller
{
    public function index(Request $request){
        if(isset($request->dtfrom)){
            $datefrForm = date("Y-m-d",strtotime(str_replace('/', '-',$request->dtfrom)));
            $datetoForm = date("Y-m-d",strtotime(str_replace('/', '-',$request->dtto)));
            $comp_code = session()->get('comp_code');
            
            // $results = DB::table('userlog')->whereBetween('datein', [$datefrForm, $datetoForm])->where('comp_code', '=', $comp_code)->get();
            $results = DB::select("SELECT * FROM userlog WHERE comp_code = '$comp_code' and DATE(datein) >= '".$datefrForm."' and DATE(datein) <= '".$datetoForm."'");
            // dd($results);
            
            return view('reports.mutasiloghist', [
                'results' => $results
            ]);
        }
        return view('reports.mutasiloghist');
    }

    public function exportExcel(Request $request)
    {
        $datefrForm = date("Y-m-d",strtotime(str_replace('/', '-',$request->dtfrom)));
        $datetoForm = date("Y-m-d",strtotime(str_replace('/', '-',$request->dtto)));
        $comp_code = session()->get('comp_code');
        $comp_name = session()->get('comp_name');

        $results = DB::select("SELECT * FROM userlog WHERE comp_code = '$comp_code' and DATE(datein) >= '".$datefrForm."' and DATE(datein) <= '".$datetoForm."'");
        
        return view('print.excel.mutasiloghist_report', compact('results', 'datefrForm', 'datetoForm', 'comp_name'));
    }

    public function exportPdf(Request $request)
    {
        $datefrForm = date("Y-m-d",strtotime(str_replace('/', '-',$request->dtfrom)));
        $datetoForm = date("Y-m-d",strtotime(str_replace('/', '-',$request->dtto)));
        $comp_code = session()->get('comp_code');
        $comp_name = session()->get('comp_name');

        $results = DB::select("SELECT * FROM userlog WHERE comp_code = '$comp_code' and DATE(datein) >= '".$datefrForm."' and DATE(datein) <= '".$datetoForm."'");
        
        return view('print.pdf.mutasiloghist_report', compact('results', 'datefrForm', 'datetoForm', 'comp_name'));
    }
}
