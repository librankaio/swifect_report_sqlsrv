<?php

namespace App\Http\Controllers;

use App\Models\User;
use Illuminate\Http\Request;
use App\Http\Requests\RegisterRequest;
use Illuminate\Support\Facades\Hash;

class RegisterController extends Controller
{
    //
    public function create()
    {
        return view('register');
    }

    public function store(Request $request)
    {
        $request->validate([
            'username' => ['required', 'string', 'min:3', 'alpha_num', 'max:25'],
            'email' => ['required', 'email'],
            'password' => ['required', 'min:8'],
            'confpassword' => ['min:8'],
        ]);
        dd($request);
        User::create([
            'username' => $request->username,
            'email' => $request->email,
            'password' => $request->password,
        ]);
        return redirect('/login');
    }
    // public function show(){
    //     return view('register');
    // }

    // public function register(RegisterRequest $request){
    //     $user = User::create($request->validated());

    //     auth()->login($user);

    //     return redirect('/')->with('Success',"Account successfully registered");
    // }
}
