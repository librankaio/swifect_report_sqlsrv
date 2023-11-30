@extends('layouts.main')
@section('content')
    <!-- Header -->
  <header>
    <div class="content_Nav">
    </div>
  </header>
  <br>
  <!-- END Header -->
  <!-- Form -->
  <form action="POST" class="container-fluid px-5 py-2">
    <div class="head-form">
      <div class="row">                
        <div class="container col-md-8 bg-white px-4 pt-3" style="border-top-left-radius: 10px; border-top-right-radius: 10px;">
          <div class="row">              
            <div class="col-md-6 bg-white">
              <div class="mb-1">
                <h5>PEMASUKKAN DOKUMEN</h5>                
              </div>              
            </div>
            <div class="col-md-6 bg-white">
              <div class="mb-3">
                {{-- <label for="exampleInputEmail1" class="form-label">Tanggal Dari</label>
                <div class="input-group date" id="datepicker">
                    <input type="text" class="form-control">
                    <span class="input-group-append">
                        <span class="input-group-text bg-white d-block">
                            <i class="fa fa-calendar"></i>
                        </span>
                    </span>
                </div> --}}
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
                <label for="exampleInputEmail1" class="form-label">Jenis Dokumen</label>
                <select class="form-select" aria-label="Default select example">
                  <option selected>All</option>
                  <option value="1">BC 2.3</option>
                  <option value="2">BC 2.6.2</option>
                  <option value="3">BC 2.7</option>
                  <option value="3">BC 4.0</option>
                </select>
                <br>
                <button type="button" class="btn btn-primary"><span> View</span></button>
                  {{-- <button type="button" class="btn btn-primary"><span> Refresh</span></button> --}}
              </div>              
            </div>
            <div class="col-md-6 bg-white">
              <div class="mb-3">
                <label for="exampleInputEmail1" class="form-label">Tanggal Dari</label>
                <div class="input-group date" id="dtfrom">
                    <input type="text" class="form-control">
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
                    <input type="text" class="form-control">
                    <span class="input-group-append">
                        <span class="input-group-text bg-white d-block">
                            <i class="fa fa-calendar"></i>
                        </span>
                    </span>
                </div>
              </div> 
            </div>
            <div class="col-md-6">
              <!-- <div class="mb-3">
                  <label for="exampleInputEmail1" class="form-label">Sampai Tanggal</label>
                  <div class="input-group date" id="datepicker">
                      <input type="text" class="form-control">
                      <span class="input-group-append">
                          <span class="input-group-text bg-white d-block">
                              <i class="fa fa-calendar"></i>
                          </span>
                      </span>
                  </div>
              </div>  -->
            </div>
          </div>
        </div> 
        <div class="container col-md-4 px-4 py-4">
          <!-- <div class="row">
            <div class="mb-3">
              <label for="exampleFormControlInput1" class="form-label">Email address</label>
              <input type="email" class="form-control" id="exampleFormControlInput1" placeholder="name@example.com">
            </div>
            <div class="mb-3">
              <label for="exampleFormControlTextarea1" class="form-label">Example textarea</label>
              <textarea class="form-control" id="exampleFormControlTextarea1" rows="3"></textarea>
            </div>
          </div> -->
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
              <button type="button" class="btn btn-success"><i class="far fa-file-excel"></i><span> Export Excel</span></button>
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
                  <label for="exampleInputEmail1" class="form-label py-2">Search :</label>
                </div>
              </div>
            </div>
            <div class="col-md-2">
              <input type="email" class="form-control" id="exampleInputEmail1" aria-describedby="emailHelp" placeholder="Search Nomor Pendaftaran...">
            </div>
          </div>
        </div>
        <table class="table table-striped table-hover">
          
          <thead>
            <tr align="center" class="" style="font-weight: bold;">
              <td scope="col" class="border-bottom-0 border-2">No</td>
              <td scope="col" class="border-bottom-0 border-2">Jenis Dokumen</td>
              <td align="center" colspan="2" class="border-bottom-0 border-2">Dokumen Pabean</td>
              <td align="center" colspan="2" class="border-bottom-0 border-2">Bukti Penerimaan Barang</td>
              <td scope="col" class="border-bottom-0 border-2">Pemasok Pengirim</td>
              <td scope="col" class="border-bottom-0 border-2">Kode Barang</td>
              <td scope="col" class="border-bottom-0 border-2">Nama Barang</td>
              <td scope="col" class="border-bottom-0 border-2">Satuan</td>
              <td scope="col" class="border-bottom-0 border-2">Jumlah</td>
              <td align="center" colspan="2" class="border-bottom-0 border-2">Nilai Barang</td>
            </tr>
            <tr align="center" class="border border-2">
              <th scope="col" class="border border-2"></th>
              <th scope="col" class="border border-2"></th>
              <th scope="col" class="border border-2">Nomor</th>
              <th scope="col" class="border border-2">Tanggal</th>
              <th scope="col" class="border border-2">Nomor</th>
              <th scope="col" class="border border-2">Tanggal</th>
              <th scope="col" class="border border-2"></th>
              <th scope="col" class="border border-2"></th>
              <th scope="col" class="border border-2"></th>
              <th scope="col" class="border border-2"></th>
              <th scope="col" class="border border-2"></th>
              <th scope="col" class="border border-2">USD</th>
              <th scope="col" class="border border-2">Rupiah</th>
            </tr>
          </thead>
          <tbody>
            <tr>                
              <th scope="row">1</th>
              <td>Mark</td>
              <td>Otto</td>
              <td>@mdo</td>
              <td>Mark</td>
              <td>Otto</td>
              <td>@mdo</td>
              <td>Mark</td>
              <td>Otto</td>
              <td>@mdo</td>
              <td>Mark</td>
              <td>Otto</td>
              <td>@mdo</td>
            </tr>
            <!-- <tr>
              <th scope="row">2</th>
              <td>Jacob</td>
              <td>Thornton</td>
              <td>@fat</td>
              <td>Mark</td>
              <td>Otto</td>
              <td>@mdo</td>
              <td>Mark</td>
              <td>Otto</td>
              <td>@mdo</td>
              <td>Mark</td>
              <td>Otto</td>
              <td>@mdo</td>
            </tr>
            <tr>
              <th scope="row">3</th>
              <td colspan="2">Larry the Bird</td>
              <td>@twitter</td>
            </tr> -->
          </tbody>
        </table>
      </div>        
      <!-- <div class="container col-md-2 bg-white ps-4 pe-4 py-4">.col</div> 
      <div class="container col-md-2 bg-white px-4 py-4">.col</div> -->
  </form>
  <!-- END Form -->
  
@endsection