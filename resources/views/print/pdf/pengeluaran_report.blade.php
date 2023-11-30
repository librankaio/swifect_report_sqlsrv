<html>
<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8">

	<title>Pengeluaran Dokumen</title>
</head>
<body class="idr" onload="window.print()">

<div style="margin-left: 0%; margin-right: 0%;">
<h5>LAPORAN PERTANGGUNG JAWABAN PENGELUARAN DOKUMEN <br>
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
                            <td scope="col" class="border-bottom-0 border-2">No</td>
                            <td scope="col" class="border-bottom-0 border-2">Jenis Dokumen</td>
                            <td align="center" colspan="2" class="border-2">Dokumen Pabean</td>
                            <td align="center" colspan="2" class="border-2">Pengeluaran Barang</td>
                            <td scope="col" class="border-bottom-0 border-2">Customer</td>
                            <td scope="col" class="border-bottom-0 border-2">Kode Barang</td>
                            <td scope="col" class="border-bottom-0 border-2">Nama Barang</td>
                            <td scope="col" class="border-bottom-0 border-2">Satuan</td>
                            <td scope="col" class="border-bottom-0 border-2">Jumlah</td>
                            <td align="center" colspan="2" class="border-2">Nilai Barang</td>
                        </tr>
                        <tr>
                            {{-- <td align="center">Nomor</td>
                            <td align="center">Tanggal</td>
                            <td align="center">Nomor</td>
                            <td align="center">Tanggal</td>
                            <td align="center">USD</td>
                            <td align="center">Rupiah</td> --}}
                            <th scope="col" class="border-top-0 border-bottom-0 border-2"></th>
                            <th scope="col" class="border-top-0 border-bottom-0  border-2"></th>
                            <th scope="col" class="border-top-0 border-bottom-0  border-2">Nomor Pendaftaran</th>
                            <th scope="col" class="border-top-0 border-bottom-0  border-2">Tanggal</th>
                            <th scope="col" class="border-top-0 border-bottom-0  border-2">Nomor</th>
                            <th scope="col" class="border-top-0 border-bottom-0  border-2">Tanggal</th>
                            <th scope="col" class="border-top-0 border-bottom-0  border-2"></th>
                            <th scope="col" class="border-top-0 border-bottom-0  border-2"></th>
                            <th scope="col" class="border-top-0 border-bottom-0  border-2"></th>
                            <th scope="col" class="border-top-0 border-bottom-0  border-2"></th>
                            <th scope="col" class="border-top-0 border-bottom-0  border-2"></th>
                            <th scope="col" class="border-top-0 border-bottom-0  border-2">Rupiah</th>
                            <th scope="col" class="border-top-0 border-bottom-0  border-2">USD</th>
                        </tr>                        
                       <tbody align="center">
              @php 
              $no=0;         
              $dpnomor = "";
              $bpbnomor = ""; @endphp     
              @isset($results)
                  {{-- @if(count($results) > 0) --}}
                  @if($no == 0)

                    @foreach ($results as $key => $item)  
                    <tr>
                          @if( $item->dpnomor == $dpnomor && $item->bpbnomor == $bpbnomor)
                            <td class="border-2"></td>
                            <td class="border-2"></td>
                            <td class="border-2"></td>
                            <td class="border-2"></td>
                            <td class="border-2"></td>
                            <td class="border-2"></td>
                            <td class="border-2"></td>
                          @elseif($item->dpnomor == $dpnomor && $item->bpbnomor != $bpbnomor)
                            <td class="border-2"></td>
                            <td class="border-2"></td>
                            <td class="border-2"></td>
                            <td class="border-2"></td>
                            <td class="border-2">{{ $item->bpbnomor }}</td>
                            <td class="border-2">{{ date("d/m/Y", strtotime($item->bpbtanggal)) }}</td>
                            <td class="border-2">{{ $item->pembeli_penerima }}</td>
                          @else
                            @php $no++ @endphp
                            <th scope="row" class="border-2">{{ $no }}</th>
                            <td class="border-2">{{ $item->jenis_dokumen }}</td>
                            <td class="border-2">{{ $item->dpnomor }}</td>                          
                            <td class="border-2">{{ date("d/m/Y", strtotime($item->dptanggal)) }}</td>
                            <td class="border-2">{{ $item->bpbnomor }}</td>
                            <td class="border-2">{{ date("d/m/Y", strtotime($item->bpbtanggal)) }}</td>
                            <td class="border-2">{{ $item->pembeli_penerima }}</td>
                        @endif
                        <td class="border-2">{{ $item->kode_barang }}</td>
                        <td class="border-2">{{ $item->nama_barang }}</td>
                        <td class="border-2">{{ $item->sat }}</td>

                        @if ($item->jumlah == 0)
                        <td class="border-2">--</td>
                        @else
                        <td class="border-2">{{ number_format($item->jumlah, 2, '.', ',') }}</td>
                        @endif
                        @if ($item->nilai_barang == 0)
                        <td class="border-2">--</td>
                        @else
                        <td class="border-2">{{ 'Rp. '.number_format($item->nilai_barang, 2, '.', ',') }}</td>
                        @endif
                        @if ($item->nilai_barang_usd == 0)
                        <td class="border-2">--</td>
                        @else
                        <td class="border-2">{{ '$. '.number_format($item->nilai_barang_usd, 2, '.', ',') }}</td>
                        @endif
                        </tr>
                        @php 
                        $dpnomor = $item->dpnomor;
                        $bpbnomor = $item->bpbnomor;
                        @endphp
                    @endforeach
                  @elseif(count($results) == 0)
                    <td colspan="13" class="border-2"> 
                      <label for="noresult" class="form-label">NO DATA RESULTS...</label>
                    </td>
                  @endif
              @endisset
            </tbody> 
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


