@extends('layouts.main')
@section('content')
<!-- Form -->
<form action="/mutasibhnbaku" method="get" class="container-fluid px-5 py-2" id="myform">
  <div class="head-form">
    <div class="row">
      <div class="container col-md-4 bg-white px-4 pt-3"
        style="border-top-left-radius: 10px; border-top-right-radius: 10px;">
        <div class="row">
          <div class="col-md-6 bg-white">
            <div class="mb-1">
              <h5>MUTASI BAHAN BAKU</h5>
            </div>
          </div>
          <div class="col-md-6 bg-white">
            <div class="mb-3">
            </div>
          </div>
        </div>
      </div>
      <div class="container col-md-4 px-4 py-4"></div>
      <div class="container col-md-4 px-4 py-4"></div>
    </div>
    <div class="row">
      <div class="container col-md-4 bg-white px-4 py-4"
        style="border-bottom-left-radius: 10px; border-bottom-right-radius: 10px;">
        <div class="row">
          <div class="col-md-6 bg-white">
            <div class="mb-3">
              <label for="exampleInputEmail1" class="form-label">Tanggal Dari</label>
              <div class="input-group date" id="dtfrom">
                <input type="text" class="form-control"
                  value="@php if(request('dtfrom')==NULL){ echo date('d/m/Y');} else{ echo $_GET['dtfrom']; } @endphp"
                  name="dtfrom">
                <span class="input-group-append">
                  <span class="input-group-text bg-white d-block">
                    <i class="fa fa-calendar"></i>
                  </span>
                </span>
              </div>
              <br>
              <button type="submit" class="btn btn-primary px-5" onclick="show_loading()"><span> View</span></button>
            </div>
          </div>
          <div class="col-md-6 bg-white">
            <div class="mb-3">
              <label for="exampleInputEmail1" class="form-label">Sampai Tanggal</label>
              <div class="input-group date" id="dtto">
                <input type="text" class="form-control"
                  value="@php if(request('dtto')==NULL){ echo date('d/m/Y');} else{ echo $_GET['dtto']; } @endphp"
                  name="dtto">
                <span class="input-group-append">
                  <span class="input-group-text bg-white d-block">
                    <i class="fa fa-calendar"></i>
                  </span>
                </span>
              </div>
            </div>
          </div>
        </div>
      </div>
      <div class="container col-md-4 px-4 py-4" style="border-bottom-right-radius: 10px;">
        {{-- <div class="row">
          <div class="col-md-6">
          </div>
        </div> --}}
      </div>
      <div class="container col-md-4 px-4 py-4">
      </div>
    </div>
  </div>
  <br>
  <div class="row">
    <div class="container col-md-12 bg-white ps-4 pe-3 py-4" style="border-radius: 10px;">
      <div class="nav-table px-1">
        <div class="row d-flex">
          <div class="col-md-6"></div>
          <div class="col-md-6 text-end">
            <button type="submit" formaction="exportpdfbhnbaku" formtarget="_blank" class="btn btn-danger"><i
                class="fa-regular fa-file-pdf"></i><span> Export PDF</span></button>
            <button type="submit" formaction="exportexcelbhnbaku" class="btn btn-success"><i
                class="far fa-file-excel"></i><span> Export Excel</span></button>
            {{-- <button type="button" class="btn btn-primary"><i class="fas fa-print"></i><span> Print</span></button>
            --}}
          </div>
        </div>
      </div>
      <div class="nav-table py-2 px-1">
        <div class="row d-flex">
          <div class="col-md-6"></div>
          <div class="col-md-2"></div>
          <div class="col-md-2">
            <div class="row">
              <div class="col-md-6"></div>
              <div class="col-md-6 text-end">
                {{-- <label for="searchtext" class="form-label py-2">Search :</label> --}}
              </div>
            </div>
          </div>
          <div class="col-md-2">
            {{-- <input type="text" class="form-control" id="searchtext" aria-describedby="searchtext" name="searchtext"
              placeholder="Search Nomor Pendaftaran..."> --}}
          </div>
        </div>
      </div>
      <div class="table-responsive">
        <table class="table table-striped table-hover" style="padding-right: 1rem;" id="datatable">
          <thead>
            <tr align="center" class="" style="font-weight: bold;">
              <th scope="col" class="border-bottom-0 border-end-0 border-2">No</th>
              <th scope="col" class="border-bottom-0 border-end-0 border-2">Kode Barang</th>
              <th align="center" class="border-bottom-0 border-end-0 border-2">Nama Barang</th>
              <th align="center" class="border-bottom-0 border-end-0 border-2">Satuan</th>
              <th scope="col" class="border-bottom-0 border-end-0 border-2">Saldo Awal</th>
              <th scope="col" class="border-bottom-0 border-end-0 border-2">Pemasukkan</th>
              <th scope="col" class="border-bottom-0 border-end-0 border-2">Pengeluaran</th>
              <th scope="col" class="border-bottom-0 border-end-0 border-2">Penyesuaian (Adjustment)</th>
              <th scope="col" class="border-bottom-0 border-2">Stock Akhir</th>
              <th align="center" class="border-bottom-0 border-end-0 border-2">Stock Opname</th>
              <th align="center" class="border-bottom-0 border-end-0 border-2">Selisih</th>
              <th align="center" class="border-bottom-0 border-2">Keterangan</th>
            </tr>
          </thead>
          <tbody>
            @php
            $no=0;
            $codemitem = "" @endphp
            @isset($results)
            {{-- @if(count($results) > 0) --}}
            @if($no == 0)
            @foreach ($results as $key => $item)
            <tr>
              @php $no++ @endphp
              <th scope="row" class="border-2">{{ $no }}</th>
              <td class="border-2">{{ $item->code_mitem }}</td>
              <td class="border-2">{{ $item->name_mitem }}</td>
              <td class="border-2">{{ $item->satuan }}</td>
              @if ($item->stock_awal == 0)
              <td class="border-2">--</td>
              @else
              <td class="border-2">{{ number_format($item->stock_awal, 2, '.', ',') }}</td>
              @endif
              @if ($item->stock_in == 0)
              <td class="border-2">--</td>
              @else
              <td class="border-2">{{ number_format($item->stock_in, 2, '.', ',') }}</td>
              @endif
              @if ($item->stock_out == 0)
              <td class="border-2">--</td>
              @else
              <td class="border-2">{{ number_format($item->stock_out, 2, '.', ',') }}</td>
              @endif
              <td class="border-2">--</td>
              @if ($item->stock_akhir == 0)
              <td class="border-2">--</td>
              @else
              <td class="border-2">{{ number_format($item->stock_akhir, 2, '.', ',') }}</td>
              @endif
              @if ($item->stock_opname == 0)
              <td class="border-2">--</td>
              @else
              <td class="border-2">{{ number_format($item->stock_opname, 2, '.', ',') }}</td>
              @endif
              <td class="border-2">--</td>
              <td class="border-2">Sesuai</td>
            </tr>
            @endforeach
            @elseif(count($results) == 0)
            <td colspan="13" class="border-2">
              <label for="noresult" class="form-label">NO DATA RESULTS...</label>
            </td>
            @endif
          </tbody>
        </table>
      </div>
      <div class="row">
        <div class="col-md-6 py-3">
          {{-- <div class="d-flex justify-content-start">
            Showing
            {{ $results->firstItem() }}
            to
            {{ $results->lastItem() }}
            of
            {{ $results->total() }}
            Entries
          </div> --}}
        </div>
        <div class="col-md-6">
          <div class="d-flex justify-content-end">
            {{-- {{ $results->appends(request()->input())->links() }} --}}
          </div>
        </div>
        @endisset
      </div>
    </div>
</form>
<!-- END Form -->
<script type="text/javascript">
  $(document).ready(function() {
    $('#datatable').dataTable({
      // "ordering":false,
      responsive: true,
       columnDefs: [
        { width: '40%', targets: 1 },
        { width: '40%', targets: 2 }
      ]
    });

    $('#datatable').css({
      'width': '100%',
      'padding-right': '0px',
    });

    // $('#datatable').DataTable();

    new DataTable('#myTable', {
    columnDefs: [{ width: '20%', targets: 0 }]
});
  });
</script>
@endsection