<template lang="">
    <div>
        <loading :active="loading" 
            :can-cancel="true"
            loader="bars" 
            color="#38b4c5"
            :height=100
            :width=200
            :on-cancel="onCancel"
            :is-full-page="true">
        </loading>
    <div class="content-header row">
    </div>
    <div class="content-body">
        <!-- Stats -->
        <div class="row">
            <div class="col-lg-4">
                <div class="card">
                    <div class="card-content">
                        <div class="media align-items-stretch">
                            <div class="p-2 text-center bg-primary bg-darken-2">
                                <i class="icon-user font-large-2 white"></i>
                            </div>
                            <div class="p-2 bg-gradient-x-primary white media-body">
                                <h3>Personas Caracterizadas</h3>
                                <h3 class="text-bold-400 mb-0"><i class="feather icon-arrow-up"></i> {{ datos.numero_femenino + datos.numero_masculino }}</h3>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-lg-4">
                <div class="card">
                    <div class="card-content">
                        <div class="media align-items-stretch">
                            <div class="p-2 text-center bg-danger bg-darken-2">
                                <i class="icon-home font-large-2 white"></i>
                            </div>
                            <div class="p-2 bg-gradient-x-danger white media-body">
                                <h3>Hogares Caracterizados</h3>
                                <h3 class="text-bold-400 mb-0"><i class="feather icon-arrow-up"></i> {{ datos.numero_hogares }}</h3>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-lg-4">
                <div class="card">
                    <div class="card-content">
                        <div class="media align-items-stretch">
                            <div class="p-2 text-center bg-warning bg-darken-2">
                                <i class="icon-globe font-large-2 white"></i>
                            </div>
                            <div class="p-2 bg-gradient-x-warning white media-body">
                                <h3>Personas Desplazadas</h3>
                                <h3 class="text-bold-400 mb-0"><i class="feather icon-arrow-up"></i>{{ datos.desplazados }}</h3>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        <!--Recent Orders & Monthly Salse -->
        <div class="row">
            <div class="col-lg-6">
                <div class="card" style="height: fit-content !important;">
                    <div class="card-header">
                        <h4 class="card-title">Piramide Poblacional</h4>
                        <a class="heading-elements-toggle"><i class="fa fa-ellipsis-v font-medium-3"></i></a>
                        <div class="heading-elements">
                            <ul class="list-inline mb-0">
                                <li><a data-action="reload"><i class="feather icon-rotate-cw"></i></a></li>
                                <li><a data-action="expand"><i class="feather icon-maximize"></i></a></li>
                            </ul>
                        </div>
                        <hr>
                    </div>
                    <div class="card-content">
                        <div class="table-responsive">
                            <div id="grafico_edad" style="height: 465px">

                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-lg-6">
                <div class="card">
                    <div class="card-content text-center">
                        <br>
                        <h4 style="margin-bottom: -0.5rem !important;" class="card-title">Personas Caracterizadas por Sexo</h4>
                        <div id="grafico_sexo" style="height: 220px">

                        </div>
                    </div>
                </div>

                <div class="card">
                    <div class="card-content text-center">
                        <div class="sales-growth-chart">
                            <br>
                            <h4 style="margin-bottom: 0.5rem !important;" class="card-title">Personas Caracterizadas por Corregimiento</h4>
                            <div id="grafico_corregimiento" style="height: 245px">

                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    </div>
</template>
<script>
import * as am4core from "@amcharts/amcharts4/core";
import * as am4charts from "@amcharts/amcharts4/charts";
import am4themes_animated from "@amcharts/amcharts4/themes/animated";
import * as dashboardService from "../../services/dashboard_service";
import Loading from 'vue3-loading-overlay';
import 'vue3-loading-overlay/dist/vue3-loading-overlay.css';

am4core.useTheme(am4themes_animated);

export default {
    components: {
        Loading
    },
    data() {
        return {
            chart_edad: null,
            chart_sexo: null,
            chart_corregimiento: null,
            datos: [],
            loading: false,
        }
    },
    mounted() {
       this.consultarDatos();
    },
    methods: {
        async consultarDatos(){
            this.loading = true;
            await dashboardService.datosDashboard().then(respuesta => {
                this.datos = respuesta.data;
                this.generarGraficoEdad();
                this.generarGraficoSexo();
                this.generarGraficoCorregimiento();
            });
            this.loading = false;
        },
        generarGraficoEdad() {
            // Create chart instance
            var chart = am4core.create("grafico_edad", am4charts.XYChart);

            // Add data
            chart.data = [];

            var maximo = 0;
            this.datos.piramide_edad.forEach(element => {
                chart.data.push({
                    "age": element[2],
                    "male": (-1) * element[0],
                    "female": element[1]
                });

                if(element[0] >= maximo){
                    maximo = element[0];
                }

                if(element[1] >= maximo){
                    maximo = element[1];
                }
            });

            // Use only absolute numbers
            chart.numberFormatter.numberFormat = "#.#s";

            // Create axes
            var categoryAxis = chart.yAxes.push(new am4charts.CategoryAxis());
            categoryAxis.dataFields.category = "age";
            categoryAxis.renderer.grid.template.location = 0;
            categoryAxis.renderer.inversed = true;
            categoryAxis.renderer.minGridDistance = 10;
            

            var valueAxis = chart.xAxes.push(new am4charts.ValueAxis());
            valueAxis.extraMin = 0.1;
            valueAxis.extraMax = 0.1;
            valueAxis.min = -1 * (maximo);
            valueAxis.max = maximo;
            valueAxis.renderer.minGridDistance = 40;
            valueAxis.renderer.ticks.template.length = 5;
            valueAxis.renderer.ticks.template.disabled = false;
            valueAxis.renderer.ticks.template.strokeOpacity = 0.4;
            valueAxis.renderer.labels.template.adapter.add("text", function(text) {
            return text == "Male" || text == "Female" ? text : text;
            })

            // Create series
            var male = chart.series.push(new am4charts.ColumnSeries());
            male.dataFields.valueX = "male";
            male.dataFields.categoryY = "age";
            male.clustered = false;

            var maleLabel = male.bullets.push(new am4charts.LabelBullet());
            maleLabel.label.text = "{valueX}";
            maleLabel.label.hideOversized = false;
            maleLabel.label.truncate = false;
            maleLabel.label.horizontalCenter = "right";
            maleLabel.label.dx = -10;

            var female = chart.series.push(new am4charts.ColumnSeries());
            female.dataFields.valueX = "female";
            female.dataFields.categoryY = "age";
            female.clustered = false;

            var femaleLabel = female.bullets.push(new am4charts.LabelBullet());
            femaleLabel.label.text = "{valueX}";
            femaleLabel.label.hideOversized = false;
            femaleLabel.label.truncate = false;
            femaleLabel.label.horizontalCenter = "left";
            femaleLabel.label.dx = 10;
           
            var maleRange = valueAxis.axisRanges.create();
            maleRange.value = -1 * (maximo);
            maleRange.endValue = 0;
            maleRange.label.text = "Masculino";
            maleRange.label.fill = chart.colors.list[0];
            maleRange.label.dy = 20;
            maleRange.label.fontWeight = '600';
            maleRange.grid.strokeOpacity = 0;
            maleRange.grid.stroke = male.stroke;

            var femaleRange = valueAxis.axisRanges.create();
            femaleRange.value = 0;
            femaleRange.endValue = maximo;
            femaleRange.label.text = "Femenino";
            femaleRange.label.fill = chart.colors.list[1];
            femaleRange.label.dy = 20;
            femaleRange.label.fontWeight = '600';
            femaleRange.grid.strokeOpacity = 0;
            femaleRange.grid.stroke = female.stroke;


            this.chart_edad = chart;
        },
        generarGraficoSexo() {
            var chart = am4core.create("grafico_sexo", am4charts.PieChart);

            // Add data
            chart.data = [
                {
                    "country": "Masculino",
                    "litres": this.datos.numero_masculino,
                    "color": am4core.color("#ff7588") // Specify the color for Masculino
                },
                {
                    "country": "Femenino",
                    "litres": this.datos.numero_femenino,
                    "color": am4core.color("#ffa87d") // Specify the color for Femenino
                }
            ];

            var pieSeries = chart.series.push(new am4charts.PieSeries());
            pieSeries.dataFields.value = "litres";
            pieSeries.dataFields.category = "country";
            pieSeries.slices.template.stroke = am4core.color("#fff");
            pieSeries.slices.template.strokeOpacity = 1;

            // Set the slice colors based on the "color" field in the data
            pieSeries.slices.template.propertyFields.fill = "color";

            pieSeries.hiddenState.properties.opacity = 1;
            pieSeries.hiddenState.properties.endAngle = -90;
            pieSeries.hiddenState.properties.startAngle = -90;

            chart.hiddenState.properties.radius = am4core.percent(0);

            this.chart_sexo = chart;
        },
        generarGraficoCorregimiento() {
            var chart = am4core.create("grafico_corregimiento", am4charts.PieChart);
            // Add data
            chart.data = [];

            this.datos.por_corregimientos.forEach(element => {
                chart.data.push({
                    "country": element.nombre,
                    "litres": element.cantidad
                });
            });
            

            var pieSeries = chart.series.push(new am4charts.PieSeries());
            pieSeries.dataFields.value = "litres";
            pieSeries.dataFields.category = "country";
            pieSeries.slices.template.stroke = am4core.color("#fff");
            pieSeries.slices.template.strokeOpacity = 1;

            pieSeries.hiddenState.properties.opacity = 1;
            pieSeries.hiddenState.properties.endAngle = -90;
            pieSeries.hiddenState.properties.startAngle = -90;

            chart.hiddenState.properties.radius = am4core.percent(0);

            this.chart_corregimiento = chart;
        },
    },
}
</script>
<style >
    i {
        line-height: 1.45;
    }
</style>