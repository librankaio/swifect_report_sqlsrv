<?php

namespace App\Http\Controllers;

use Carbon\Carbon;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Auth;
use Illuminate\Support\Facades\DB;

class UserController extends Controller
{
    public function getLogin(Request $request){
        $username = $request->get('email');
        $password = $request->get('password');
        if(filter_var($username, FILTER_VALIDATE_EMAIL)) {
            //user sent their email
            Auth::attempt(['email' => $username, 'password' => $password]);
        } else {
            //they sent their username instead
            Auth::attempt(['name' => $username, 'password' => $password]);
        }

       //was any of those correct ?
        if ( Auth::check()){                        
            $current_date_time = Carbon::now()->toDateTimeString();
            DB::table('userlog')->insert(['username' =>Auth::user()->name, 'tbl'=>'ONLINE', 'idtbl'=> '0', 'notbl'=>'', 'act'=>'LOGIN', 'comp_code'=>'BYC', 'usin'=>1,'datein'=>$current_date_time]);
            //send them where they are going
            return redirect()->route('dashboard');
        }
            return redirect()->back();
    }
}
