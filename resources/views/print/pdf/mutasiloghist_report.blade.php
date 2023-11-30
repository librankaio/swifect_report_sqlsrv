<html>

<head>

    <title>MUTASI LOG HISTORY</title>
</head>

<body class="idr" onload="window.print()">
    <div style="margin-left: 0%; margin-right: 0%;">
        <h5>LAPORAN PERTANGGUNG JAWABAN MUTASI LOG HISTORY<br>
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
                    <td scope="col" class="border-bottom-0 border-end-0 border-2">Nama User</td>
                    <td scope="col" class="border-bottom-0 border-end-0 border-2">Jenis Transaksi</td>
                    <td align="center" class="border-bottom-0 border-end-0 border-2">Action</td>
                    <td align="center" class="border-bottom-0 border-end-0 border-2">Tanggal/Waktu</td>
                </tr>
                @if(count($results) > 0)
                @foreach ($results as $key => $item)
                <tr>
                    <th scope="row" class="border-2">{{ $item->username }}</th>
                    <td class="border-2">{{ $item->tbl }}</td>
                    <td class="border-2">{{ $item->act }}</td>
                    <td class="border-2">{{ $item->datein }}</td>
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