
function generateBarData(companies){

const counts = {};

companies.forEach(c => {

let category = c.company_based_on;

if(!category || category.trim()===""){
category = "Not Available";
}

category = category.trim();

if(!counts[category]){
counts[category] = 0;
}

counts[category]++;

});

return Object.keys(counts).map(k => ({
key:k,
value:counts[k]
}));

}

const barLastUpdated = Date.now();

function calculateBarPercentages(items){
const total = items.reduce((sum,i)=> sum + i.value ,0);

return items.map(i=>({
...i,
percentage: total>0 ? Math.round((i.value/total)*100) : 0
}));
}

function sortBarData(items){
return items.slice().sort((a,b)=> b.value - a.value);
}
let horizontalChart;
function initHorizontalChart(){

fetch("companies.json")
.then(res => res.json())
.then(companies => {

const counts = {};

companies.forEach(c => {

let category = c.company_based_on;

if(!category || category.trim()===""){
category = "Not Available";
}

category = category.trim();

if(!counts[category]){
counts[category] = 0;
}

counts[category]++;

});

const barChartData = Object.keys(counts).map(k => ({
key:k,
value:counts[k]
}));

horizontalChart = echarts.init(document.getElementById("chartContainer"));

const sort = sortBarData(barChartData);
const seriesData = calculateBarPercentages(sort);


const option = {

backgroundColor:'#e7e9f1',

title:{
text:'📊 Based On ',
left:'left',
top:10,
textStyle:{
color:'#101010',
fontSize:18,
fontWeight:'bold'
}
},

graphic:{
elements:[{
type:'text',
right:'20',
top:'10',
style:{
text:`Updated: ${new Date(barLastUpdated).toLocaleTimeString()}`,
fontSize:12,
fill:'#aaa'
}
}]
},

tooltip:{
trigger:'axis',
axisPointer:{type:'shadow'},
backgroundColor:'#f4f6f8',
borderColor:'#0e0f0f',
textStyle:{color:'#161515'},
formatter:function(params){
const p=params[0];
const pct=seriesData.find(i=> i.key===p.name).percentage;
return `<b>${p.name}</b><br>Value: ${p.value}<br>Percent: ${pct}%`;
}
},

grid:{
left:'25%',
right:'10%',
top:'25%',
bottom:'10%'
},

dataZoom:[
{
type:'slider',
yAxisIndex:0,
start:0,
end:40
},
{
type:'inside',
yAxisIndex:0,
start:0,
end:40
}
],
xAxis:{
type:'value',
splitLine:{show:false},
axisLabel:{color:'#4a4747'}
},

yAxis:{
type:'category',
data:seriesData.map(i=> i.key),
inverse:true,
axisLine:{show:false},
axisTick:{show:false},
axisLabel:{
color:'#0e0d0d',
fontSize:13
}
},

series:[{
type:'bar',
barWidth:'55%',
data:seriesData.map(i=>({
value:i.value,
itemStyle:{
borderRadius:[0,10,10,0],
color:new echarts.graphic.LinearGradient(0,0,1,0,[
{offset:0,color:'#3b82f6'},
{offset:1,color:'#22c55e'}
]),
shadowBlur:10,
shadowColor:'rgba(0,0,0,0.4)'
}
})),

label:{
show:true,
position:'right',
color:'#0b0b0b',
fontWeight:'bold',
formatter:function(params){
const pct=seriesData.find(i=> i.key===params.name).percentage;
return pct+'%';
}
}

}]

};

horizontalChart.setOption(option);

window.addEventListener('resize',()=>horizontalChart.resize());
});

}
document.addEventListener("DOMContentLoaded",initHorizontalChart);

function updateHorizontalChart(companies){

const counts = {};

companies.forEach(c => {

let category = c.company_based_on;

if(!category || category.trim()===""){
category = "Not Available";
}

category = category.trim();

if(!counts[category]){
counts[category] = 0;
}

counts[category]++;

});

const barChartData = Object.keys(counts).map(k => ({
key:k,
value:counts[k]
}));

const sort = sortBarData(barChartData);
const seriesData = calculateBarPercentages(sort);

const option = {

yAxis:{
data:seriesData.map(i=> i.key)
},

series:[{
data:seriesData.map(i=>({
value:i.value,
itemStyle:{
borderRadius:[0,10,10,0],
color:new echarts.graphic.LinearGradient(0,0,1,0,[
{offset:0,color:'#3b82f6'},
{offset:1,color:'#22c55e'}
]),
shadowBlur:10,
shadowColor:'rgba(0,0,0,0.4)'
}
}))
}]

};

horizontalChart.setOption(option);

}