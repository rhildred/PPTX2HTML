var prettifyXml = function(xml, tab){
    var formatted = '', indent= '';
    tab = tab || '\t';
    xml.split(/>\s*</).forEach(function(node) {
        if (node.match( /^\/\w/ )) indent = indent.substring(tab.length); // decrease indent by one 'tab'
        formatted += indent + '<' + node + '>\r\n';
        if (node.match( /^<?\w[^>]*[^\/]$/ )) indent += tab;              // increase indent
    });
    return formatted.substring(1, formatted.length-3);
}


$(document).ready(function() {

	if (window.Worker) {
		
		var $result = $("#result");
		var isDone = false;
		
		$("#uploadBtn").on("change", function(evt) {
			
			isDone = false;
			
			$result.html("");
			$("#load-progress").text("0%").attr("aria-valuenow", 0).css("width", "0%");
			$("#result_block").removeClass("hidden").addClass("show");
			
			var fileName = evt.target.files[0];
			
			// Read the file
			var reader = new FileReader();
			reader.onload = (function(theFile) {
				return function(e) {
					
					// Web Worker
					var worker = new Worker('./js/worker.js');
				
					worker.addEventListener('message', function(e) {
						
						var msg = e.data;
						
						switch(msg.type) {
							case "progress-update":
								$("#load-progress").text(msg.data.toFixed(2) + "%")
									.attr("aria-valuenow", msg.data.toFixed(2))
									.css("width", msg.data.toFixed(2) + "%");
								break;
							case "slide":
								$result.append(msg.data);
								break;
							case "processMsgQueue":
								processMsgQueue(msg.data);
								break;
							case "pptx-thumb":
								$("#pptx-thumb").attr("src", "data:image/jpeg;base64," + msg.data);
								break;
							case "slideSize":
								if (localStorage) {
									localStorage.setItem("slideWidth", msg.data.width);
									localStorage.setItem("slideHeight", msg.data.height);
								} else {
									alert("Browser don't support Web Storage!");
								}
								break;
							case "globalCSS":
								$result.append("<style>" + msg.data + "</style>");
								break;
							case "ExecutionTime":
								$("#info_block").html("Execution Time: " + msg.data + " (ms)");
								isDone = true;
								worker.postMessage({
									"type": "getMsgQueue"
								});
								break;
							case "WARN":
								console.warn('Worker: ', msg.data);
								break;
							case "ERROR":
								console.error('Worker: ', msg.data);
								$("#error_block").text(msg.data);
								break;
							case "DEBUG":
								console.debug('Worker: ', msg.data);
								break;
							case "INFO":
							default:
								console.info('Worker: ', msg.data);
								//$("#info_block").html($("#info_block").html() + "<br><br>" + msg.data);
						}
						
					}, false);
					
					worker.postMessage({
						"type": "processPPTX",
						"data": e.target.result
					});
					
				}
			})(fileName);
			reader.readAsArrayBuffer(fileName);
			
		});
						
		$("#download-reveal-btn").click(function () {
			if (!isDone) { return; }
			var bodyHtml = prettifyXml($result.html());
			var html = 
`<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="utf-8">
<title>Demo</title>
<link rel="stylesheet" href="https://rhildred.github.io/PPTX2HTML/reveal/css/reveal.css">
<link rel="stylesheet" href="https://rhildred.github.io/PPTX2HTML/reveal/css/theme/pptx.css" id="theme">  

<!-- Code syntax highlighting -->
<link rel="stylesheet" href="https://rhildred.github.io/PPTX2HTML/reveal/lib/css/zenburn.css">
<link rel="stylesheet" href="https://rhildred.github.io/PPTX2HTML/reveal/plugin/chart/nv.d3.min.css">

<!--Add support for earlier versions of Internet Explorer -->
<!--[if lt IE 9]>
<script src="lib/js/html5shiv.js"></script>
<![endif]-->
<style>
section {
	width: 100%;
	height: 690px;
	position: relative;
	text-align: center;
	overflow: hidden;
}

section div.block {
	position: absolute;
	top: 0px;
	left: 0px;
	width: 100%;
}

section div.content {
	display: flex;
	flex-direction: column;
	/*
	justify-content: center;
	align-items: flex-end;
	*/
}

section div.v-up {
	justify-content: flex-start;
}
section div.v-mid {
	justify-content: center;
}
section div.v-down {
	justify-content: flex-end;
}

section div.h-left {
	align-items: flex-start;
	text-align: left;
}
section div.h-mid {
	align-items: center;
	text-align: center;
}
section div.h-right {
	align-items: flex-end;
	text-align: right;
}

section div.up-left {
	justify-content: flex-start;
	align-items: flex-start;
	text-align: left;
}
section div.up-center {
	justify-content: flex-start;
	align-items: center;
}
section div.up-right {
	justify-content: flex-start;
	align-items: flex-end;
}
section div.center-left {
	justify-content: center;
	align-items: flex-start;
	text-align: left;
}
section div.center-center {
	justify-content: center;
	align-items: center;
}
section div.center-right {
	justify-content: center;
	align-items: flex-end;
}
section div.down-left {
	justify-content: flex-end;
	align-items: flex-start;
	text-align: left;
}
section div.down-center {
	justify-content: flex-end;
	align-items: center;
}
section div.down-right {
	justify-content: flex-end;
	align-items: flex-end;
}

section span.text-block {
	/* display: inline-block; */
}

li.slide {
	margin: 10px 0px;
	font-size: 18px;
}

div.footer {
	text-align: center;
}

section table {
	position: absolute;
}

section table, section th, section td {
	border: 1px solid black;
}

section svg.drawing {
	position: absolute;
	overflow: visible;
}

section img {
	all: initial;
}
</style>;
</head>
<body>
<div class="reveal">
<div id='slides' class='slides'>
${bodyHtml}
</div>
</div>
<script src="https://rhildred.github.io/PPTX2HTML/reveal/lib/js/head.min.js"></script>
<script src="https://rhildred.github.io/PPTX2HTML/reveal/js/reveal.js"></script>
<script src="https://rhildred.github.io/PPTX2HTML/reveal/plugin/chart/d3.min.js"></script>
<script src="https://rhildred.github.io/PPTX2HTML/reveal/plugin/chart/nv.d3.min.js"></script>
<script>
Reveal.initialize({
	width: +localStorage.getItem("slideWidth"),
	height: +localStorage.getItem("slideHeight"),
	controls: true,
	progress: true,
	history: true,
	center: true,
	transition: 'convex',
	dependencies: [{
		src: 'https://rhildred.github.io/PPTX2HTML/reveal/plugin/highlight/highlight.js',
		async: true,
		condition: function() { return !!document.querySelector('pre code'); },
		callback: function() { hljs.initHighlightingOnLoad(); }
	}]
});
</script>
</body>
</html>`;
			var blob = new Blob([html], {type: "text/html;charset=utf-8"});
			saveAs(blob, "slides.html");
		});
		
		$("#to-reveal-btn").click(function () {
			if (localStorage) {
				localStorage.setItem("slides", LZString.compressToUTF16($result.html()));
				window.open("./reveal/demo.html", "_blank");
			} else {
				alert("Browser don't support Web Storage!");
			}
		});
		
	} else {
		
		alert("Browser don't support Web Worker!");
		
	}
	
});

function processMsgQueue(queue) {
	for (var i=0; i<queue.length; i++) {
		processSingleMsg(queue[i].data);
	}
}

function processSingleMsg(d) {
	
	var chartID = d.chartID;
	var chartType = d.chartType;
	var chartData = d.chartData;

	var data =  [];
	
	var chart = null;
	switch (chartType) {
		case "lineChart":
			data = chartData;
			chart = nv.models.lineChart()
						.useInteractiveGuideline(true);
			chart.xAxis.tickFormat(function(d) { return chartData[0].xlabels[d] || d; });
			break;
		case "barChart":
			data = chartData;
			chart = nv.models.multiBarChart();
			chart.xAxis.tickFormat(function(d) { return chartData[0].xlabels[d] || d; });
			break;
		case "pieChart":
		case "pie3DChart":
			data = chartData[0].values;
			chart = nv.models.pieChart();
			break;
		case "areaChart":
			data = chartData;
			chart = nv.models.stackedAreaChart()
						.clipEdge(true)
						.useInteractiveGuideline(true);
			chart.xAxis.tickFormat(function(d) { return chartData[0].xlabels[d] || d; });
			break;
		case "scatterChart":
			
			for (var i=0; i<chartData.length; i++) {
				var arr = [];
				for (var j=0; j<chartData[i].length; j++) {
					arr.push({x: j, y: chartData[i][j]});
				}
				data.push({key: 'data' + (i + 1), values: arr});
			}
			
			//data = chartData;
			chart = nv.models.scatterChart()
						.showDistX(true)
						.showDistY(true)
						.color(d3.scale.category10().range());
			chart.xAxis.axisLabel('X').tickFormat(d3.format('.02f'));
			chart.yAxis.axisLabel('Y').tickFormat(d3.format('.02f'));
			break;
		default:
	}
	
	if (chart !== null) {
		
		d3.select("#" + chartID)
			.append("svg")
			.datum(data)
			.transition().duration(500)
			.call(chart);
		
		nv.utils.windowResize(chart.update);
		
	}
	
}
