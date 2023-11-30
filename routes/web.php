<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\LoginController;
use App\Http\Controllers\MutasiBhnBakuController;
use App\Http\Controllers\MutasiBrgJadiController;
use App\Http\Controllers\MutasiLogHistController;
use App\Http\Controllers\MutasiMesinController;
use App\Http\Controllers\MutasiScrapController;
use App\Http\Controllers\MutasiWinProcessController;
use App\Http\Controllers\PemasukkanController;
use App\Http\Controllers\RegisterController;
use App\Http\Controllers\PengeluaranController;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Auth;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

// Route::get('register', [RegisterController::class, 'create'])->name('register');

// Route::post('register', [RegisterController::class, 'store'])->name('register');



Route::get('/', [LoginController::class, 'index'])->name('login');
Route::post('/', [LoginController::class, 'postLogin'])->name('postlogin');
Route::get('logout', [LoginController::class, 'logout'])->name('logout');
Route::group(['middleware' => 'guest'], function () {

});


Route::group(['middleware' => ['auth']], function () {

    //Pemasukkan
    Route::get('pemasukan', [PemasukkanController::class, 'index'])->name('pemasukan');
    Route::get('exportexcelpemasukan', [PemasukkanController::class, 'exportExcel'])->name('exportexcel');
    // Route::get('exportpdfpemasukan', [PemasukkanController::class, 'exportPdf'])->name('exportpdf');
    Route::get('exportpdfpemasukan', [PemasukkanController::class, 'exportPdf'])->name('exportpdf');
    
    //Pengeluaran
    Route::get('pengeluaran', [PengeluaranController::class, 'index'])->name('pengeluaran');
    Route::get('exportexcelpengeluaran', [PengeluaranController::class, 'exportExcel'])->name('exportexcel');
    Route::get('exportpdfpengeluaran', [PengeluaranController::class, 'exportPdf'])->name('exportpdf');
    // Route::get('searchp', [PengeluaranController::class, 'searchPengeluaran'])->name('searchpengeluaran');
    
    //MutasiBahanBaku
    Route::get('mutasibhnbaku', [MutasiBhnBakuController::class, 'index'])->name('mutasibhnbaku');
    Route::get('exportexcelbhnbaku', [MutasiBhnBakuController::class, 'exportExcel'])->name('exportexcel');
    Route::get('exportpdfbhnbaku', [MutasiBhnBakuController::class, 'exportPdf'])->name('exportpdf');
    
    //MutasiBarangJAdi
    Route::get('mutasibrgjadi', [MutasiBrgJadiController::class, 'index'])->name('mutasibrgjadi');
    Route::get('exportexcelbrgjadi', [MutasiBrgJadiController::class, 'exportExcel'])->name('exportexcel');
    Route::get('exportpdfbrgjadi', [MutasiBrgJadiController::class, 'exportPdf'])->name('exportPdf');

    // MutasiScrap
    Route::get('mutasiscrap', [MutasiScrapController::class, 'index'])->name('mutasiscrap');
    Route::get('exportexcelscrap', [MutasiScrapController::class, 'exportExcel'])->name('exportexcel');
    Route::get('exportpdfscrap', [MutasiScrapController::class, 'exportPdf'])->name('exportpdf');

    // MutasiMesin
    Route::get('mutasimesin', [MutasiMesinController::class, 'index'])->name('mutasimesin');
    Route::get('exportexcelmutasimesin', [MutasiMesinController::class, 'exportExcel'])->name('exportexcel');
    Route::get('exportpdfmutasimesin', [MutasiMesinController::class, 'exportPdf'])->name('exportpdf');

    // MutasiWorkInProcess
    Route::get('mutasiwip', [MutasiWinProcessController::class, 'index'])->name('mutasiwip');
    Route::get('exportexcelmutasiwip', [MutasiWinProcessController::class, 'exportExcel'])->name('exportexcel');
    Route::get('exportpdfmutasiwip', [MutasiWinProcessController::class, 'exportPdf'])->name('exportpdf');

    // MutasiLogHist
    Route::get('mutasiloghist', [MutasiLogHistController::class, 'index'])->name('mutasiloghist');
    Route::get('exportexcelmutasiloghist', [MutasiLogHistController::class, 'exportExcel'])->name('exportexcel');
    Route::get('exportpdfmutasiloghist', [MutasiLogHistController::class, 'exportPdf'])->name('exportpdf');
});
