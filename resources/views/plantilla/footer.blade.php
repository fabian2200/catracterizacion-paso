 <!-- BEGIN: Footer-->
 <footer class="footer fixed-bottom footer-dark navbar-shadow">
    <p style="max-width: 100% !important;" class="clearfix blue-grey lighten-2 text-sm-center mb-0 px-2 container center-layout"><span
            class="float-md-left d-block d-md-inline-block">Copyright &copy; {{ date('Y') }} <a
                class="text-bold-800 grey darken-2" href="https://www.leeringenieria.com/Nosotros.php"
                target="_blank">LEER INGENIERIA S.A.S </a></span><span class="float-md-right d-none d-lg-block">CARACTERIZACIÓN POBLACION AFRODESCENDIENTES <i style="color: #ffff" class="fa fa-users"></i></span></p>
</footer>
<!-- END: Footer-->

<script src="{{ asset('js/jquery.3.6.min.js') }}"></script>

<script src="{{ asset('js/app.js') }}"></script>

<!-- BEGIN: Vendor JS-->
<script src="{{ asset('app-assets/vendors/js/vendors.min.js') }}"></script>
<!-- BEGIN Vendor JS-->

<!-- BEGIN: Page Vendor JS-->
<script src="{{ asset('app-assets/vendors/js/ui/jquery.sticky.js') }}"></script>
<script src="{{ asset('app-assets/vendors/js/charts/jquery.sparkline.min.js') }}"></script>
<!-- END: Page Vendor JS-->

<!-- BEGIN: Theme JS-->
<script src="{{ asset('app-assets/js/core/app-menu.js') }}"></script>
<script src="{{ asset('app-assets/js/core/app.js') }}"></script>
<!-- END: Theme JS-->

<!-- BEGIN: Page JS-->
<script src="{{ asset('app-assets/js/scripts/ui/breadcrumbs-with-stats.js') }}"></script>
<script src="{{ asset('app-assets/vendors/js/extensions/toastr.min.js') }}"></script>
<script src="{{ asset('app-assets/js/scripts/extensions/toastr.js') }}"></script>
<script src="{{ asset('app-assets/js/scripts/tooltip/tooltip.js') }}"></script>
<script src="{{ asset('app-assets/js/html2pdf.bundle.min.js') }}"></script>
<script src="{{ asset('app-assets/js/printThis.js') }}"></script>
<!-- END: Page JS-->

<script src="{{ asset('app-assets/vendors/js/tables/datatable/datatables.min.js') }}"></script>


<script>
    toastr.options = {
        "closeButton": true,
        "newestOnTop": false,
        "progressBar": false,
        "positionClass": "toast-top-right",
        "preventDuplicates": false,
        "onclick": null,
        "showDuration": "500",
        "hideDuration": "1000",
        "timeOut": "5000",
        "extendedTimeOut": "1000",
        "showEasing": "swing",
        "hideEasing": "linear",
        "showMethod": "fadeIn",
        "hideMethod": "fadeOut"
    }

</script>