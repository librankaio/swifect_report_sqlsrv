@extends('layouts.main')
@section('content')
  <!-- Form -->
  <form action="/mutasiloghist" method="get" class="container-fluid px-5 py-2" id="myform">
    <div class="head-form">
      <div class="row">                
        <div class="container col-md-4 bg-white px-4 pt-3" style="border-top-left-radius: 10px; border-top-right-radius: 10px;">
          <div class="row">              
            <div class="col-md-6 bg-white">
              <div class="mb-1">
                <h5>MUTASI LOG HISTORY</h5>                
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
        <div class="container col-md-4 bg-white px-4 py-4" style="border-bottom-left-radius: 10px; border-bottom-right-radius: 10px;">
          <div class="row">              
            <div class="col-md-6 bg-white">
              <div class="mb-3">   
                <label for="exampleInputEmail1" class="form-label">Tanggal Dari</label>
                <div class="input-group date" id="dtfrom">
                    <input type="text" class="form-control"  value="@php if(request('dtfrom')==NULL){ echo date('d/m/Y');} else{ echo $_GET['dtfrom']; } @endphp"  name="dtfrom">
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
                    <input type="text" class="form-control" value="@php if(request('dtto')==NULL){ echo date('d/m/Y');} else{ echo $_GET['dtto']; } @endphp" name="dtto">
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
              <button type="submit" class="btn btn-danger" formaction="exportpdfmutasiloghist" formtarget="_blank"><i class="fas fa-print"></i><span> Print</span></button>
                <button type="submit" formaction="exportexcelmutasiloghist" class="btn btn-success"><i class="far fa-file-excel"></i><span> Export Excel</span></button>              
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
              {{-- <input type="text" class="form-control" id="searchtext" aria-describedby="searchtext" name="searchtext" placeholder="Search Nomor Pendaftaran..."> --}}
            </div>
          </div>
        </div>
        <table class="table table-striped table-hover" id="datatable" style="padding-right: 0.5em;">
          
          <thead>
            <tr align="center" class="" style="font-weight: bold;">
              <td scope="col" class="border-bottom-0 border-end-0 border-2">Nama User</td>
              <td scope="col" class="border-bottom-0 border-end-0 border-2">Jenis Transaksi</td>
              <td align="center" class="border-bottom-0 border-end-0 border-2">Action</td>
              <td align="center" class="border-bottom-0 border-end-0 border-2">Tanggal/Waktu</td>
            </tr>
          </thead>
          <tbody>   
            @isset($results)
                @foreach ($results as $key => $item)  
                <tr>
                  <th scope="row" class="border-2">{{ $item->username }}</th>
                  <td class="border-2">{{ $item->tbl }}</td>
                  <td class="border-2">{{ $item->act }}</td>
                  <td class="border-2">{{ $item->datein }}</td>
                </tr>
                @endforeach
                {{-- <td colspan="13" class="border-2"> 
                  <label for="noresult" class="form-label">NO DATA RESULTS...</label>
                </td> --}}
            @endisset
          </tbody>
        </table>
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
          {{-- @endisset --}}
        </div>
      </div>        
  </form>
  <!-- END Form -->
<script type="text/javascript">
  $(document).ready(function() {
    // $('#datatable').dataTable(
    //         {"ordering":false});

    $('#datatable').DataTable();
  });
</script>
@endsection