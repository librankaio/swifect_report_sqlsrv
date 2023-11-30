@extends('layouts.main')
@section('topscript')
  <style type="text/css">
    .dataTables_filter {
      display: none;
    }
  </style>
@stop
@section('content')
@php auth()->user() @endphp
  <!-- Form -->
  {{-- <form action="{{ route('searchpengeluaran') }}" method="get" class="container-fluid px-5 py-2"> --}}
  <form action="/pengeluaran" method="get" class="container-fluid px-5 py-2">
    <div class="head-form">
      <div class="row">                
        <div class="container col-md-8 bg-white px-4 pt-3" style="border-top-left-radius: 10px; border-top-right-radius: 10px;">
          <div class="row">              
            <div class="col-md-6 bg-white">
              <div class="mb-1">
                <h5>PENGELUARAN DOKUMEN</h5>                
              </div>              
            </div>
            <div class="col-md-6 bg-white">
              <div class="mb-3">
              </div>
            </div>
          </div>    
        </div>        
        
        <div class="container col-md-4 px-4 py-4"></div>           
      </div>
      <div class="row">                
        <div class="container col-md-4 bg-white px-4 py-4" style="border-bottom-left-radius: 10px;">
          <div class="row">              
            <div class="col-md-6 bg-white">
              <div class="mb-3">                
                <label for="jenisdokumen" class="form-label">Jenis Dokumen</label>
                <select class="form-select" aria-label="Default select example" name="jenisdok">
                  <?php 
                  if(request()->input('jenisdok') == NULL){ 
                  ?>
                    <option value='All'>All</option>
                    <option value="BC 2.5">BC 2.5</option>
                    <option value="BC 2.6.1">BC 2.6.1</option>
                    <option value="BC 2.7">BC 2.7</option>
                    <option value="BC 3.0">BC 3.0</option>
                    <option value="BC 4.1">BC 4.1</option>
                  <?php }else{ ?>
                    <option selected value='{{ $_GET['jenisdok'] }}'>{{ $_GET['jenisdok'] }}</option>
                    <option value='All'>All</option>
                    <option value="BC 2.5">BC 2.5</option>
                    <option value="BC 2.6.1">BC 2.6.1</option>
                    <option value="BC 2.7">BC 2.7</option>
                    <option value="BC 3.0">BC 3.0</option>
                    <option value="BC 4.1">BC 4.1</option>
                  <?php } ?>
                </select>
                <br>
                  <button type="submit" class="btn btn-primary px-5" onclick="show_loading()"><span> View</span></button>
                  {{-- <button type="button" class="btn btn-primary"><span> Refresh</span></button> --}}
              </div>              
            </div>
            <div class="col-md-6 bg-white">
              <div class="mb-3">
                <label for="exampleInputEmail1" class="form-label">Tanggal Dari</label>
                <div class="input-group date" id="dtfrom">
                    <input type="text" class="form-control"  value="@php if(request()->input('dtfrom')==NULL){ echo date('d/m/Y');} else{ echo $_GET['dtfrom']; } @endphp"  name="dtfrom">
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
        <div class="container col-md-4 bg-white px-4 py-4" style="border-bottom-right-radius: 10px;">
          <div class="row">
            <div class="col-md-6 bg-white">
              <div class="mb-3">
                <label for="exampleInputEmail1" class="form-label">Sampai Tanggal</label>
                <div class="input-group date" id="dtto">
                    <input type="text" class="form-control" value="@php if(request()->input('dtto')==NULL){ echo date('d/m/Y');} else{ echo $_GET['dtto']; } @endphp" name="dtto">
                    <span class="input-group-append">
                        <span class="input-group-text bg-white d-block">
                            <i class="fa fa-calendar"></i>
                        </span>
                    </span>
                </div>
              </div> 
            </div>
            <div class="col-md-6">
            </div>
          </div>
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
                <button type="submit" formaction="exportpdfpengeluaran" formtarget="_blank" class="btn btn-danger"><i class="fa-regular fa-file-pdf"></i><span> Export PDF</span></button>
                <button type="submit" formaction="exportexcelpengeluaran" class="btn btn-success"><i class="far fa-file-excel"></i><span> Export Excel</span></button> 
              {{-- <button type="button" class="btn btn-primary"><i class="fas fa-print"></i><span> Print</span></button> --}}
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
                  <label for="searchtext" class="form-label py-2">Search :</label>
                </div>
              </div>
            </div>
            <div class="col-md-2">
              <input type="text" class="form-control" id="searchtext" aria-describedby="searchtext" name="searchtext" placeholder="Search Nomor Pendaftaran...">
            </div>
          </div>
        </div>
        {{-- RESPONSIVE TABLE IN GENERAL MONITOR >1400px --}}
        <div class="d-xl-none d-xxl-block">
          <table class="table table-striped table-hover px-2" id="datatable_xxl">
            <thead>
              <tr align="center" class="" style="font-weight: bold;">
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
              <tr align="center" class="border-top-0 border-2">
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
            </thead>          
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
        </div>
        {{-- END SPONSIVE TABLE IN GENERAL MONITOR --}}
        {{-- RESPONSIVE IN XL only show in >1200px --}}
      <div class="d-none d-xl-block d-xxl-none">
          <table class="table table-striped table-hover datatable_xl" id="datatable_xl" style="max-width: 950px;">
            <thead style="font-size: 11px;">
              <tr align="center" class="" style="font-weight: bold;">
                <td scope="col" class="border-bottom-0 border-2" style="width: 1px;">No</td>
                <td scope="col" class="border-bottom-0 border-2 border-start-0">Jenis Dokumen</td>
                <td align="center" colspan="2" class="border-2 border-start-0">Dokumen Pabean</td>
                <td align="center" colspan="2" class="border-2 border-start-0">Pengeluaran Barang</td>
                <td scope="col" class="border-bottom-0 border-2 border-start-0">Customer</td>
                <td scope="col" class="border-bottom-0 border-2 border-start-0">Kode Barang</td>
                <td scope="col" class="border-bottom-0 border-2 border-start-0">Nama Barang</td>
                <td scope="col" class="border-bottom-0 border-2 border-start-0">Satuan</td>
                <td scope="col" class="border-bottom-0 border-2 border-start-0">Jumlah</td>
                <td align="center" colspan="2" class="border-2 border-start-0">Nilai Barang</td>
              </tr>
              <tr align="center">
                <th scope="col" class="border-top-0 border-bottom-0 border-2" style="width: 1px;"></th>
                <th scope="col" class="border-top-0 border-bottom-0  border-2 border-start-0"></th>
                <th scope="col" class="border-top-0 border-bottom-0  border-2 border-start-0">Nomor Pendaftaran</th>
                <th scope="col" class="border-top-0 border-bottom-0  border-2 border-start-0">Tanggal</th>
                <th scope="col" class="border-top-0 border-bottom-0  border-2 border-start-0">Nomor</th>
                <th scope="col" class="border-top-0 border-bottom-0  border-2 border-start-0">Tanggal</th>
                <th scope="col" class="border-top-0 border-bottom-0  border-2 border-start-0"></th>
                <th scope="col" class="border-top-0 border-bottom-0  border-2 border-start-0"></th>
                <th scope="col" class="border-top-0 border-bottom-0  border-2 border-start-0"></th>
                <th scope="col" class="border-top-0 border-bottom-0  border-2 border-start-0"></th>
                <th scope="col" class="border-top-0 border-bottom-0  border-2 border-start-0"></th>
                <th scope="col" class="border-top-0 border-bottom-0  border-2 border-start-0">Rupiah</th>
                <th scope="col" class="border-top-0 border-bottom-0  border-2 border-start-0">USD</th>
              </tr>
            </thead>          
            <tbody style="font-size: 9px;">
              @php 
              $no=0;         
              $dpnomor = ""; @endphp        
              @isset($results)
                  {{-- @if(count($results) > 0) --}}
                  @if($no == 0)

                    @foreach ($results as $key => $item)  
                    <tr>
                          @if( $item->dpnomor == $dpnomor)
                            <td class="border-2" style="width: 1px;"></td>
                            <td class="border-2"></td>
                            <td class="border-2"></td>
                            <td class="border-2"></td>
                            <td class="border-2"></td>
                            <td class="border-2"></td>
                            <td class="border-2"></td>
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
                        $dpnomor = $item->dpnomor
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
      </div>
      {{-- END RESPONSIVE TABLE IN XL --}}
        <div class="row">
          <div class="col-md-6 py-3">
            <div class="d-flex justify-content-start">
              {{-- Showing
              {{ $results->firstItem() }}
              to
              {{ $results->lastItem() }}
              of
              {{ $results->total() }}
              Entries --}}
            </div>
          </div>
          <div class="col-md-6">
            <div class="d-flex justify-content-end">
              {{-- {{ $results->appends(request()->input())->links() }} --}}
            </div>
          </div>
          {{-- @endisset --}}
        </div>
      </div>  
  </form>
  <!-- END Form -->  
<script type="text/javascript">
  $(document).ready(function() {
    $('#datatable_xl').dataTable(
            {"ordering":false});

    $('#datatable_xl').DataTable();

    $('#datatable_xxl').dataTable(
            {"ordering":false});

    $('#datatable_xxl').DataTable();
  });
</script>
@endsection