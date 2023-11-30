<input type="checkbox" id="check">


<!-- Sidebar -->
<div class="sidebar" id="sidebar">

  <div class="head-sidebar">
    <a href="{{ 'pemasukan' }}" class="logo-btn"><img src="img/logo_pjt.png" alt="" class="logo color-light"><span>IT
        INVENTORY BC</span></img></a>
    <label for="check">
      <i class="fas fa-bars" id="sidebar_toggle"></i>
    </label>
  </div>
  <center class="profile_form">
    @php
    $imgname = session('comp_code');
    @endphp
    <img src='{{ asset("../img/".$imgname.".png") }}' alt="" class="profile_img"></img>
    {{-- <img src='/img/{{ $imgname }}.png' alt="" class="profile_img"></img> --}}
    <h4>{{ session('comp_name') }}</h4>
  </center>
  <a href="{{ 'pemasukan' }}" class="btn_Nav {{ 'pemasukan' == request()->path() ? 'btn_NavActive' : '' }}"><i
      class="fas fa-sign-in-alt"></i><span>Pemasukan Dokumen</span></a>
  <a href="{{ 'pengeluaran' }}" class="btn_Nav {{ 'pengeluaran' == request()->path() ? 'btn_NavActive' : '' }}"><i
      class="fas fa-sign-out-alt"></i><span>Pengeluaran Dokumen</span></a>
  <a href="{{ 'mutasibhnbaku' }}" class="btn_Nav {{ 'mutasibhnbaku' == request()->path() ? 'btn_NavActive' : '' }}"><i
      class="fa-regular fa-note-sticky"></i><span>Mutasi Bahan Baku</span></a>
  <a href="{{ 'mutasibrgjadi' }}" class="btn_Nav {{ 'mutasibrgjadi' == request()->path() ? 'btn_NavActive' : '' }}"><i
      class="fa-regular fa-note-sticky"></i><span>Mutasi Barang Jadi</span></a>
  <a href="{{ 'mutasiscrap' }}" class="btn_Nav {{ 'mutasiscrap' == request()->path() ? 'btn_NavActive' : '' }}"><i
      class="fa-regular fa-note-sticky"></i><span>Mutasi Scrap</span></a>
  {{-- Batas --}}
  <a href="{{ 'mutasimesin' }}" class="btn_Nav {{ 'mutasimesin' == request()->path() ? 'btn_NavActive' : '' }}"><i
      class="fa-regular fa-note-sticky"></i></i><span>Mutasi Mesin</span></a>
  <a href="{{ 'mutasiwip' }}" class="btn_Nav {{ 'mutasiwip' == request()->path() ? 'btn_NavActive' : '' }}"><i
      class="fa-regular fa-note-sticky"></i><span>Mutasi Working Process</span></a>
  <a href="{{ 'mutasiloghist' }}" class="btn_Nav {{ 'mutasiloghist' == request()->path() ? 'btn_NavActive' : '' }}"><i
      class="fa-regular fa-note-sticky"></i><span>Mutasi Log History</span></a>
  {{-- <a href="#" class="btn_Nav"><i class="fas fa-desktop"></i><span>Mutasi Bahan Baku</span></a>
  <a href="#" class="btn_Nav"><i class="fas fa-desktop"></i><span>Mutasi Barang Jadi</span></a>
  <a href="#" class="btn_Nav"><i class="fas fa-desktop"></i><span>Mutasi Mesin</span></a>
  <a href="#" class="btn_Nav"><i class="fas fa-desktop"></i><span>Mutasi Working Process</span></a>
  <a href="#" class="btn_Nav"><i class="fas fa-desktop"></i><span>Log History</span></a> --}}
</div>
<!-- END Sidebar -->