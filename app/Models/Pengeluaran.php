<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Pengeluaran extends Model
{
    use HasFactory;

    protected $table='pengeluaran_dokumen';
    protected $fillable = [
        'id',
        'jenis_dokumen',
        'nomoraju',
        'bpbnomor',
        'bpbtanggal',
        'pembeli_penerima',
        'kode_barang',
        'nama_barang',
        'sat',
        'jumlah',
        'nilai_barang',
        'nilai_barang_usd',
        'comp_code',
        'note',
        'tstatus'];
}
