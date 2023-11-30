<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>Switaatpajak | ERP APPS</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous">
    <link href="/css/style.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta2/css/all.min.css">
</head>
<body>
    
    {{-- Login Panel --}}
    {{-- <div class="container-fluid bg-primary" style="height: 1000px, width: 100%">ewq
        <div class="container bg-info text-center">asdffas</div>
    </div> --}}
    <form action="{{ route('register') }}" method="post">
        <section class="vh-100 bg-img">
            <div class="container py-5 h-100">
                <div class="row d-flex justify-content-center align-items-center h-100">
                <div class="col-12 col-md-8 col-lg-6 col-xl-5">
                    <div class="card shadow-2-strong" style="border-radius: 1rem;">
                    <div class="card-body p-5 text-center">
                        <input type="hidden" name="_token" value="{{ csrf_token() }}"/>
                        <h3 class="mb-5">Register</h3>                        
                        <div class="form-outline mb-3 text-start">
                            <label class="form-label" for="username">Username</label>
                            <input type="text" id="username" class="form-control form-control-lg" placeholder="Username" name="username" required="required" id="username" value="{{ old('username') }}" autofocus>
                            </div>
                            {{-- @if ($errors->has('username'))
                                <span class="text-danger text-left">{{ $errors->first('username') }}</span>
                            @endif --}}
                        <div class="form-outline mb-3 text-start">
                        <label class="form-label" for="email">Email</label>
                        <input type="text" id="email" class="form-control form-control-lg" placeholder="Email" name="email" required="required" id="email" value="{{ old('email') }}" autofocus/>
                        </div>
                        @if ($errors->has('email'))
                            <span class="text-danger text-left">{{ $errors->first('email') }}</span>
                        @endif

                        <div class="form-outline mb-4 text-start">
                        <label class="form-label" for="password">Password</label>
                        <input type="password" id="password" class="form-control form-control-lg" placeholder="Password" name="password" value="{{ old('password') }}" required="required">
                        </div>
                        @if ($errors->has('password'))
                            <span class="text-danger text-left">{{ $errors->first('password') }}</span>
                        @endif

                        {{-- <div class="form-outline mb-4 text-start">
                            <label class="form-label" for="confpassword">Confirm Password</label>
                            <input type="password" id="confpassword" class="form-control form-control-lg" placeholder="Confirm Password" name="confpassword" required="required" value="{{ old('confpassword') }}">
                            </div>
                            @if ($errors->has('confpassword'))
                            <span class="text-danger text-left">{{ $errors->first('confpassword') }}</span>
                        @endif --}}
                        <!-- Checkbox -->
                        {{-- <div class="form-check d-flex justify-content-start mb-4">
                        <input
                            class="form-check-input"
                            type="checkbox"
                            value=""
                            id="form1Example3"
                        />
                        <label class="form-check-label" for="form1Example3"> Remember password </label>
                        </div> --}}
                        <button class="btn btn-primary btn-lg btn-block" style="width:100%;" type="submit">Register</button>
                        <div class="py-2"></div>
                        <hr>
                        <div class="login-form pt-2">
                            Already have an account?<a href="{{ '/' }}">Login Here</a>
                        </div>
                    </div>
                    </div>
                </div>
                </div>
            </div>
        </section>
        @include('partials.copy')
    </form>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
</body>
</html>