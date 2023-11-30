<html>
<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8">

	<title>Pemasukan Dokumen</title>
</head>
<body class="idr" onload="window.print()">

<div style="margin-left: 0%; margin-right: 0%;">
<h5>LAPORAN PERTANGGUNG JAWABAN PEMASUKAN DOKUMEN <br>
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
                            <td align="center" rowspan="2">No</td>
                            <td align="center" rowspan="2">Jenis Dokumen</td>
                            <td align="center" colspan="2">Dokumen Pabean</td>
                            <td align="center" colspan="2">Bukti Penerimaan Barang</td>
                            <td align="center" rowspan="2">Pemasok Pengirim</td>
                            <td align="center" rowspan="2">Kode Barang</td>
                            <td width="30px" align="center" rowspan="2">Nama Barang</td>
                            <td align="center" rowspan="2">Satuan</td>
                            <td align="center" rowspan="2">Jumlah</td>
                            <td align="center" colspan="2">Nilai Barang</td>
                        </tr>
                        <tr>
                            <td align="center">Nomor</td>
                            <td align="center">Tanggal</td>
                            <td align="center">Nomor</td>
                            <td align="center">Tanggal</td>
                            <td align="center">USD</td>
                            <td align="center">Rupiah</td>
                        </tr>                        
                        <?php
                            $no = 0;
                            $nomoraju="";
                        ?>
                        @foreach($results as $item)
                        <tr>
                        <?php
                        if ($no==0)
                        {
                            $no++;
                        ?>   
                        <td><div style="width: 50px; word-wrap: break-word;" >{{ $no }}</div></td>
                            <td><div style="width: 50px; word-wrap: break-word;" >{{ $item->jenis_dokumen }}</div></td>
                            <td><div style="width: 50px; word-wrap: break-word;" >{{ $item->dpnomor }}</div></td>
                            <td><div style="width: 50px; word-wrap: break-word;" >{{ date("d/m/Y", strtotime($item->dptanggal))  }}</div></td>
                            <td><div style="width: 60px; word-wrap: break-word; " >{{ $item->bpbnomor }}</div></td>
                            <td><div style="width: 50px; word-wrap: break-word;" >@if(date("d/m/Y", strtotime($item->bpbtanggal)) != '30/11/-0001')
                                {{ date("d/m/Y", strtotime($item->bpbtanggal)) }}
                                @else
                                @endif</div></td>
                            <td><div style="width: 100px; word-wrap: break-word;" >{{ $item->pemasok_pengirim }}</div></td>
                        <?php    
                        }
                        else
                        {
                                if ( $nomoraju == $item->nomoraju )
                                {
                                ?>
                                    <td></td>
                                    <td></td>
                                    <td></td>
                                    <td></td>
                                    <td></td>
                                    <td></td>
                                    <td></td>
                                <?php
                                }
                                else
                                {
                                    $no++;
                                ?>    
                                    <td><div style="width: 50px; word-wrap: break-word;" >{{ $no }}</div></td>
                                    <td><div style="width: 50px; word-wrap: break-word;" >{{ $item->jenis_dokumen }}</div></td>
                                    <td><div style="width: 50px; word-wrap: break-word;" >{{ $item->dpnomor }}</div></td>
                                    <td><div style="width: 50px; word-wrap: break-word;" >{{ date("d/m/Y", strtotime($item->dptanggal))  }}</div></td>
                                    <td><div style="width: 60px; word-wrap: break-word; " >{{ $item->bpbnomor }}</div></td>
                                    <td><div style="width: 50px; word-wrap: break-word;" >@if(date("d/m/Y", strtotime($item->bpbtanggal)) != '30/11/-0001')
                                        {{ date("d/m/Y", strtotime($item->bpbtanggal)) }}
                                        @else
                                        @endif</div></td>
                                    <td><div style="width: 100px; word-wrap: break-word;" >{{ $item->pemasok_pengirim }}</div></td>
                                <?php
                                }
                        }
                        ?>
                        
                            <td><div style="width: 50px; word-wrap: break-word;" >{{ $item->kode_barang }}</div></td>
                            <td><div style="width: 100px; word-wrap: break-word;" >{{ $item->nama_barang }}</div></td>
                            <td><div style="width: 50px; word-wrap: break-word;" >{{ $item->sat }}</div></td>
                            <td><div style="width: 50px; word-wrap: break-word;" >{{ number_format($item->jumlah, 2, '.', ',') }}</div></td>
                            <td><div style="width: 100px; word-wrap: break-word;" >$ {{ number_format($item->nilai_barang_usd, 2, '.', ',') }}</div></td>
                            <td><div style="width: 100px; word-wrap: break-word;" >Rp. {{ number_format($item->nilai_barang, 2, '.', ',') }}</div></td>
                        </tr>
                        <?php
                        $nomoraju=$item->nomoraju;
                        ?>
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


