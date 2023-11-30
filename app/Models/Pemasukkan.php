<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Pemasukkan extends Model
{
    use HasFactory;

    protected $table='pemasukan_dokumen';
    protected $fillable = [
        'id',
        'jenis_dokumen',
        'nomoraju',
        'dpbnomor',
        'dpbtanggal',
        'bpbnomor',
        'bpbtanggal',
        'pemasok_pengirim',
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
