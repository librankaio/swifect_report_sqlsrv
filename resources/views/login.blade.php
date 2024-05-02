<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>SWIBC APPS</title>
    {{-- <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous"> --}}
    <link rel="stylesheet" href="{{ asset("../bootstrap/css/bootstrap.min.css") }}">
    <link href="/css/style.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta2/css/all.min.css">
</head>
<body>
    
    {{-- Login Panel --}}
    {{-- <div class="container-fluid bg-primary" style="height: 1000px, width: 100%">ewq
        <div class="container bg-info text-center">asdffas</div>
    </div> --}}
    <form action="{{ route('login') }}" method="post">
        <section class="vh-100 bg-img">
            <div class="container py-5 h-100">
                <div class="row d-flex justify-content-center align-items-center h-100">
                    <div class="col-12 col-md-8 col-lg-6 col-xl-5">
                        <img src="/img/logo-swifect-logo.png" alt=""></img>
                        <h1 style="font-weight: bold; color: #2B58AF">IT INVENTORY ONLINE BC</h1>
                    </div>
                    <div class="col-12 col-md-8 col-lg-6 col-xl-5">
                        <div class="card shadow-2-strong" style="border-radius: 1rem;">
                        <div class="card-body p-5 text-center">
                            <input type="hidden" name="_token" value="{{ csrf_token() }}" />
                            {{-- <h3 class="mb-5">IT INVENTORY ONLINE BC</h3> --}}
                
                            <div class="form-outline mb-3 text-start">
                            <label class="form-label" for="username">Username</label>
                            <input type="text" id="username" class="form-control form-control-lg" placeholder="Username" name="username" required="required" autofocus/>
                            </div>
                
                            <div class="form-outline mb-3 text-start">
                            <label class="form-label" for="typePasswordX-2">Password</label>
                            <input type="password" id="typePasswordX-2" class="form-control form-control-lg" placeholder="Password" name="password" required="required" autofocus/>
                            </div>
                
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
                            <br>
                            <button class="btn btn-primary btn-lg btn-block" style="width:100%;" type="submit">Login</button>
                            <div class="py-2"></div>
                            <hr>
                            {{-- <div class="login-form pt-2">
                                Don't have an account?<a href="{{ 'register' }}">Register Here</a>
                            </div> --}}
                            <div class="login-form pt-2">
                                <a href="https://swifect.com/" target="_blank">PT.Swifect Solusi Indonesia 2022 ©</a>
                            </div>
                            <div class="login-form pt-2">
                                Go to Swifect Repository page click <a href="https://repository.swiapps.com/" target="_blank">Here</a>
                            </div>
                        </div>
                        </div>
                    </div>                    
                </div>
            </div>
        </section>
    </form>
    {{-- <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script> --}}
    <script src="{{ asset("../bootstrap/js/bootstrap.bundle.min.js") }}"></script>
</body>
</html>