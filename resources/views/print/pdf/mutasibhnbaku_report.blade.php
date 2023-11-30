<html>
<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8">

	<title>Mutasi Bahan Baku Dokumen</title>
</head>
<body class="idr" onload="window.print()">

<div style="margin-left: 0%; margin-right: 0%;">
<h5>LAPORAN PERTANGGUNG JAWABAN MUTASI BAHAN BAKU <br>
{{-- <?php echo $itype." ".$iname; ?> <br> --}}
<?php
    if ($datefrForm == NULL AND $datetoForm == NULL) {
    }else{
?>
PERIODE {{ $datefrForm }} S.D {{ $datetoForm }}</h5>
<?php } ?>
<center>
<table id="mytable" border="1px" cellspacing="0">
       <tr>
                            <td align="center">No</td>
                            <td align="center">Kode Barang</td>
                            <td align="center">Nama Barang</td>
                            <td align="center">Satuan</td>
                            <td align="center">Saldo Awal</td>
                            <td align="center">Pemasukkan</td>
                            <td align="center">Pengeluaran</td>
                            <td align="center">Penyesuaian (Adjustment)</td>
                            <td align="center">Stock Akhir</td>
                            <td align="center">Stock Opname</td>
                            <td align="center">Selisih</td>
                            <td align="center">Keterangan</td>
                        </tr>      
                        <?php
                            $no = 0;
                        ?>
                        @foreach($results as $item)
                        <tr>
                            @php
                            $no++;
                            @endphp
                                    <td><div style="word-wrap: break-word;" >{{ $no }}</div></td>
                                    <td><div style="word-wrap: break-word;" >{{ $item->code_mitem }}</div></td>
                                    <td><div style="word-wrap: break-word;" >{{ $item->name_mitem }}</div></td>
                                    <td><div style="word-wrap: break-word;" >{{ $item->satuan }}</div></td>
                                    <td><div style="width: 60px; word-wrap: break-word; " >{{ number_format($item->stock_awal, 2, '.', ',') }}</div></td>
                                    <td><div style="width: 60px; word-wrap: break-word; " >{{ number_format($item->stock_in, 2, '.', ',') }}</div></td>
                                    <td><div style="width: 60px; word-wrap: break-word; " >{{ number_format($item->stock_out, 2, '.', ',') }}</div></td>
                                    <td><div style="word-wrap: break-word;" >0</div></td>
                                    <td><div style="width: 60px; word-wrap: break-word; " >{{ number_format($item->stock_akhir, 2, '.', ',') }}</div></td>
                                    <td><div style="word-wrap: break-word;" >--</div></td>
                                    @if ($item->stock_opname == 0)
                                    <td><div style="word-wrap: break-word;" >--</div></td>
                                    @else
                                    <td><div style="word-wrap: break-word;" >{{ number_format($item->stock_opname, 2, '.', ',') }}</div></td>
                                    @endif
                                    <td><div style="word-wrap: break-word;" >Sesuai</div></td>
                        @endforeach
    </table>
</center>
<br><br>
<p>
	<footer><a href="http://www.swifect.com">~ Swifect Custom Application ~</a></footer>
</div>
</body>
</html>

<style type="text/css" media="print">
  @page { size: landscape; margin: 0px auto; }
</style>


