document.addEventListener("DOMContentLoaded", loadData);

function loadData(){

fetch("companies.json")
.then(res => res.json())
.then(companies => {

console.log(companies[0]);

const counts = {};

companies.forEach(c => {

const name = (c.company_name || "")
.toLowerCase()
.replace(/\./g,"")
.replace(/\s+/g," ")
.trim();

let type = "Other";

if(/pvt\s*ltd$/.test(name) || name.endsWith("private limited")){
    type = "Pvt Ltd";
}
else if(name.endsWith(" llp")){
    type = "LLP";
}
else if(
    (name.endsWith(" ltd") || name.endsWith(" limited")) &&
    !name.includes("pvt")
){
   type = "Ltd";
}

counts[type] = (counts[type] || 0) + 1;

});

const rawData = {
KPIs: Object.keys(counts).map(k => ({
StartTimeUTC: k,
OEE1: counts[k]
}))
};

drawChart(rawData);

});
}

function drawChart(rawData){

const lineData = rawData.KPIs;

const margin={top:20,right:30,bottom:40,left:50};

const container = document.querySelector(".viz");

const containerWidth = container.clientWidth;

const width = containerWidth - margin.left - margin.right;
const height = 350 - margin.top - margin.bottom;

d3.select(".viz").selectAll("*").remove(); // clear old chart

const svg=d3.select(".viz")
.append("svg")
.attr("viewBox",`0 0 ${width+margin.left+margin.right} ${height+margin.top+margin.bottom}`)
.attr("preserveAspectRatio","xMidYMid meet");

const group=svg.append("g")
.attr("transform",`translate(${margin.left},${margin.top})`);

const xScale=d3.scalePoint()
.domain(lineData.map(d=>d.StartTimeUTC))
.range([0,width]);

const yScale=d3.scaleLinear()
.domain([0, d3.max(lineData, d => d.OEE1) * 1.1])
.range([height,0]);

// gradient
const defs=svg.append("defs");

const gradient=defs.append("linearGradient")
.attr("id","line-gradient");

gradient.append("stop")
.attr("offset","0%")
.attr("stop-color","#22c55e");

gradient.append("stop")
.attr("offset","100%")
.attr("stop-color","#3b82f6");

// line
const line=d3.line()
.x(d=>xScale(d.StartTimeUTC))
.y(d=>yScale(d.OEE1))
.curve(d3.curveMonotoneX);

const path=group.append("path")
.datum(lineData)
.attr("fill","none")
.attr("stroke","url(#line-gradient)")
.attr("stroke-width",4)
.attr("d",line);

// animation
const length=path.node().getTotalLength();

path
.attr("stroke-dasharray",length)
.attr("stroke-dashoffset",length)
.transition()
.duration(1500)
.attr("stroke-dashoffset",0);

// tooltip
const tooltip=d3.select(".tooltip");

// points
group.selectAll("circle")
.data(lineData)
.enter()
.append("circle")
.attr("cx",d=>xScale(d.StartTimeUTC))
.attr("cy",d=>yScale(d.OEE1))
.attr("r",5)
.attr("fill","#252424")
.on("mouseover",(event,d)=>{
tooltip.style("opacity",1)
.html(`Value: ${d.OEE1}`);
})
.on("mousemove",(event)=>{
tooltip
.style("left",(event.pageX+10)+"px")
.style("top",(event.pageY-20)+"px");
})
.on("mouseout",()=>{
tooltip.style("opacity",0);
});

// axes
group.append("g")
.attr("transform",`translate(0,${height})`)
.call(d3.axisBottom(xScale));

group.append("g")
.call(d3.axisLeft(yScale));

}

// redraw when window resizes
window.addEventListener("resize", loadData);

function updateLineChart(companies){

const counts = {};

companies.forEach(c => {

const name = (c.company_name || "")
.toLowerCase()
.replace(/\./g,"")
.replace(/\s+/g," ")
.trim();

let type = "Other";

if(/pvt\s*ltd$/.test(name) || name.endsWith("private limited")){
type = "Pvt Ltd";
}
else if(name.endsWith(" llp")){
type = "LLP";
}
else if(
(name.endsWith(" ltd") || name.endsWith(" limited")) &&
!name.includes("pvt")
){
type = "Ltd";
}

counts[type] = (counts[type] || 0) + 1;

});

const rawData = {
KPIs: Object.keys(counts).map(k => ({
StartTimeUTC: k,
OEE1: counts[k]
}))
};

drawChart(rawData);

}