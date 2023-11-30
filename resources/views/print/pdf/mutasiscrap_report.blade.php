<html>

<head>

    <title>Mutasi SCRAP</title>
</head>

<body class="idr" onload="window.print()">
    <div style="margin-left: 0%; margin-right: 0%;">
        <h5>LAPORAN PERTANGGUNG JAWABAN MUTASI SCRAP<br>
            {{-- Nama Companny --}}
            {{ $comp_name }}<br>
            <?php
    if ($datefrForm == NULL AND $datetoForm == NULL) {
    }else{
        $datefr = strtotime($datefrForm);
        $datefrFormat = date('d/m/Y',$datefr);
        $dateto = strtotime($datetoForm);
        $datetoFormat = date('d/m/Y',$dateto);   
?>
            PERIODE {{ $datefrFormat }} S.D {{ $datetoFormat }}
        </h5>
        <?php } ?>

        <center>
            <table id="mytable" border="1px" cellspacing="0">
                <tr>
                    <td scope="col" class="border-bottom-0 border-end-0 border-2">No</td>
                    <td scope="col" class="border-bottom-0 border-end-0 border-2">Kode Barang</td>
                    <td align="center" class="border-bottom-0 border-end-0 border-2">Nama Barang</td>
                    <td align="center" class="border-bottom-0 border-end-0 border-2">Satuan</td>
                    <td scope="col" class="border-bottom-0 border-end-0 border-2">Saldo Awal</td>
                    <td scope="col" class="border-bottom-0 border-end-0 border-2">Pemasukkan</td>
                    <td scope="col" class="border-bottom-0 border-end-0 border-2">Pengeluaran</td>
                    <td scope="col" class="border-bottom-0 border-end-0 border-2">Penyesuaian (Adjustment)</td>
                    <td scope="col" class="border-bottom-0 border-2">Stock Akhir</td>
                    <td align="center" class="border-bottom-0 border-end-0 border-2">Stock Opname</td>
                    <td align="center" class="border-bottom-0 border-end-0 border-2">Selisih</td>
                    <td align="center" class="border-bottom-0 border-2">Keterangan</td>
                </tr>
                {{-- <tr>
                    <td align="center" scope="col" class="border-top-0 border-bottom-0  border-2">Nomor</td>
                    <td align="center" scope="col" class="border-top-0 border-bottom-0  border-2">Tanggal</td>
                    <td align="center" scope="col" class="border-top-0 border-bottom-0  border-2">Nomor</td>
                    <td align="center" scope="col" class="border-top-0 border-bottom-0  border-2">Tanggal</td>
                    <td align="center" scope="col" class="border-top-0 border-bottom-0  border-2">USD</td>
                    <td align="center" scope="col" class="border-top-0 border-bottom-0  border-2">Rupiah</td>
                </tr> --}}
                @if(count($results) > 0)
                @php $no=0; @endphp
                @foreach ($results as $key => $item)
                <tr>
                    @php
                    $no++;
                    @endphp
                    <td scope="row" class="border-2">{{ $no }}</td>
                    <td class="border-2">{{ $item->code_mitem }}</td>
                    <td class="border-2">{{ $item->name_mitem }}</td>
                    <td class="border-2">{{ $item->satuan }}</td>
                    @if ($item->stock_awal == 0)
                    <td class="border-2">0</td>
                    @else
                    <td class="border-2">{{ number_format($item->stock_awal, 2, '.', ',') }}</td>
                    @endif
                    @if ($item->stock_in == 0)
                    <td class="border-2">0</td>
                    @else
                    <td class="border-2">{{ number_format($item->stock_in, 2, '.', ',') }}</td>
                    @endif
                    @if ($item->stock_out == 0)
                    <td class="border-2">0</td>
                    @else
                    <td class="border-2">{{ number_format($item->stock_out, 2, '.', ',') }}</td>
                    @endif
                    <td class="border-2">0</td>
                    @if ($item->stock_akhir == 0)
                    <td class="border-2">0</td>
                    @else
                    <td class="border-2">{{ number_format($item->stock_akhir, 2, '.', ',') }}</td>
                    @endif
                    @if ($item->stock_opname == 0)
                    <td class="border-2">0</td>
                    @else
                    <td class="border-2">{{ number_format($item->stock_opname, 2, '.', ',') }}</td>
                    @endif
                    <td class="border-2">0</td>
                    <td class="border-2">Sesuai</td>
                </tr>
                @endforeach
                @elseif(count($results) == 0)
                <td colspan="13" class="border-2">
                    <label for="noresult" class="form-label">NO DATA RESULTS...</label>
                </td>
                @endif
            </table>
        </center>
        <br><br>
        <p>
        <footer><a href="http://www.swifect.com">~ Swifect Inventory BC ~</a></footer>
    </div>
</body>

</html>

{{-- <style type="text/css" media="print">
    @page {
        size: landscape;
        margin: 0px auto;
    }
</style>

<style type="text/css" media="print">
    @page {
        margin: 0px auto;
    }
</style> --}}

<style type="text/css" media="print">
    @page { size: landscape; margin: 0px auto; }
  </style>