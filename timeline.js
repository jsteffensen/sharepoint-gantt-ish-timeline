/*********************************************************
*
*  dependencies: jQuery, d3js, underscorejs, bluebirdjs
*
**********************************************************/
// item version history are at: http://sharepoint/.../_layouts/15/Versions.aspx?list=[list id]&ID=[item id]&IsDlg=1


var margin = 40;
var width = 2400;
var height = 1500;
var mwidth = width - 2 * margin;
var mheight = height - 2 * margin;
var charttitle = 'My Timeline';

var startArrivalDate = new Date();

// Overwrite today's date to match sample data - DELETE THIS ////////////////////////////////////////////////////////////////////////////
startArrivalDate = new Date('2020-03-12');
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// hard-code date range to 15 months for best visualization
// for dynamic range see createXaxisData()
var chartStartDate = new Date();
chartStartDate.setMonth(startArrivalDate.getMonth()-1);
var endArrivalDate = new Date();
endArrivalDate.setMonth(endArrivalDate.getMonth()+14);

var xstart = 0;
var xstop = 0;

$(document).ready(function() {

    init();
});

function init() {

    getMilestoneData().then(function(data) {
      return getCalendarData(data);
    }).then(function(data) {
      return createXaxisData(data);
    }).then(function(data) {
      return createYaxisData(data);
    }).then(function(data) {
      return createDiagram(data);
    }).then(function(data) {
      //console.log(data);
    }, function(e) {
      console.log(e);
    });

}

/*function getMilestoneData() {

  return new Promise(function(resolve, reject) {

    urlstring = 'http://sharepoint/_vti_bin/listdata.svc/Calendar?$orderby=ParentId';

      $.ajax({
        url: urlstring,
        type: 'GET',
        dataType: 'json',
        success: function(data) {
          calendar = {};
          
          calendar.milestones = {};
          calendar.titles = {};

          // create object to read milestone values from

          for(i=0; i<data.d.results.length; i++) {
              key = 'id_'+ data.d.results[i].Id;
              title = data.d.results[i].Title;
              calendar.titles[key] = title;
            if(data.d.results[i].ParentId == 0) {
              value = parseInt(data.d.results[i].Arrival.replace('/Date(', '').replace(')/', ''));
              calendar.milestones[key] = value;
            } else if(data.d.results[i].ParentId > 0) {
              parentKey = 'id_'+ data.d.results[i].ParentId;
              offsetMillis = Math.ceil( parseInt(data.d.results[i].OffsetDays) * (1000 * 60 * 60 * 24));
              value = calendar.milestones[parentKey] + offsetMillis;
              calendar.milestones[key] = value;
            }

          }
         
          resolve(calendar);
        },
        error: function(jqxhr, status, exception) {
          reject(exception);
        }
      });

  });
}*/

// use sample data
function getMilestoneData() {

  return new Promise(function(resolve, reject) {

          calendar = {};
          
          calendar.milestones = {};
          calendar.titles = {};

          for(i=0; i<sharepointData.length; i++) {
              key = 'id_'+ sharepointData[i].Id;
              title = sharepointData[i].Title;
              calendar.titles[key] = title;
            if(sharepointData[i].ParentId == 0) {
              value = parseInt(sharepointData[i].Arrival.replace('/Date(', '').replace(')/', ''));
              calendar.milestones[key] = value;
            } else if(sharepointData[i].ParentId > 0) {
              parentKey = 'id_'+ sharepointData[i].ParentId;
              offsetMillis = Math.ceil( parseInt(sharepointData[i].OffsetDays) * (1000 * 60 * 60 * 24));
              value = calendar.milestones[parentKey] + offsetMillis;
              calendar.milestones[key] = value;
            }
          }
         
          resolve(calendar);

  });
}

// get data from sharepoint Calendar
/*function getCalendarData(calendar) {

  return new Promise(function(resolve, reject) {

    // To use filter directly in the listdata.svc request:
    //startdate = new Date();
    //enddate = new Date();
    //enddate.setMonth(enddate.getMonth()+14);
    //stringStart = isoString(startdate);
    //stringEnd = isoString(enddate);
    //filterFormat = "?$filter=(({0} gt datetime'{1}') and ({2} lt datetime'{3}'))",
    //filterQuery = String.format(filterFormat, 'Arrival', stringStart, 'EndTime', stringEnd),
    //urlstring = 'http://sharepoint/_vti_bin/listdata.svc/Calendar' + filterQuery + '&$orderby=Arrival';

    urlstring = 'http://sharepoint/_vti_bin/listdata.svc/Calendar?$orderby=Arrival';

      $.ajax({
        url: urlstring,
        type: 'GET',
        dataType: 'json',
        success: function(data) {

          calendar.data = [];

          for(i=0; i<data.d.results.length; i++) {

            arrival = calendar.milestones['id_' + data.d.results[i].Id]
            if(arrival > startArrivalDate && arrival < endArrivalDate) {

              if(data.d.results[i].ParentId == 0) {

                calendar.data.push(data.d.results[i]);

              } else if(data.d.results[i].ParentId > 0) {

                data.d.results[i].Dependent = true;
                milestoneValue = calendar.milestones['id_' + data.d.results[i].ParentId];
                milestoneDate = new Date(milestoneValue);
                dependentArrival = milestoneDate.setTime(milestoneDate.getTime() + (parseInt(data.d.results[i].OffsetDays) * 86400000));

                data.d.results[i].Arrival = '/Date(' + dependentArrival + ')/';
                calendar.data.push(data.d.results[i]);
              } else {
                console.log('Error: ParentId not set.');
              }
            }
          }

          // sort to achieve staggered visualization
          calendar.data = _.sortBy(calendar.data, 'Arrival');

          resolve(calendar);
        },
        error: function(jqxhr, status, exception) {
          reject(exception);
        }
      });

  });
}*/

// use sample data
function getCalendarData(calendar) {

  return new Promise(function(resolve, reject) {

          calendar.data = [];

          for(i=0; i<sharepointData.length; i++) {

            arrival = calendar.milestones['id_' + sharepointData[i].Id]
            if(arrival > startArrivalDate && arrival < endArrivalDate) {

              if(sharepointData[i].ParentId == 0) {

                calendar.data.push(sharepointData[i]);

              } else if(sharepointData[i].ParentId > 0) {

                sharepointData[i].Dependent = true;
                milestoneValue = calendar.milestones['id_' + sharepointData[i].ParentId];
                milestoneDate = new Date(milestoneValue);
                dependentArrival = milestoneDate.setTime(milestoneDate.getTime() + (parseInt(sharepointData[i].OffsetDays) * 86400000));

                sharepointData[i].Arrival = '/Date(' + dependentArrival + ')/';
                calendar.data.push(sharepointData[i]);
              } else {
                console.log('Error: ParentId not set.');
              }
            }
          }

          calendar.data = _.sortBy(calendar.data, 'Arrival');

          resolve(calendar);


  });
}

function createXaxisData(calendar) {

  return new Promise(function(resolve, reject) {

      calendar.xaxis = [];
      minDate = 0;
      maxDate = 0;

      // find highest and lowest date values to determine x axis range dynamically ...
      for(i=0; i<calendar.data.length; i++) {

	arrival = parseInt(calendar.data[i].Arrival.replace('/Date(', '').replace(')/', ''));
	endTime = parseInt(calendar.data[i].EndTime.replace('/Date(', '').replace(')/', ''));

        if(minDate == 0 || arrival<minDate) {
          minDate = arrival;
        }

        if(maxDate == 0 || endTime>maxDate) {
          maxDate = endTime;
        }

      }

      // ... or use the hard-coded range
      calendar.xaxis.push(chartStartDate);
      calendar.xaxis.push(endArrivalDate);

      resolve(calendar);

  });
}

function createYaxisData(calendar) {

  return new Promise(function(resolve, reject) {

      calendar.yaxis = [];

      for(i=0; i<calendar.data.length; i++) {
        if( !_.contains(calendar.yaxis, calendar.data[i].Team) ) {
          calendar.yaxis.push(calendar.data[i].Team);
        }
      }
      calendar.yaxis = _.sortBy(calendar.yaxis);
      calendar.yaxis.reverse();
      resolve(calendar);

  });
}

function createDiagram(calendar) {

  return new Promise(function(resolve, reject) {

      svg = d3.select('svg')
        .attr('width', width)
        .attr('height', height)
        .attr('style', 'outline: thin solid #ddd; fill; #cccccc');

      // clear if it already holds data
      svg.selectAll('*').remove();

      // append a group to hold all the diagram elements
      chart = svg.append('g')
        .attr('transform', 'translate(' + margin + ', ' + margin + ')')
        .attr('id', 'chartarea');

      xscale = d3.scaleTime()
        .range([0, mwidth])
        .domain(calendar.xaxis);

      yscale = d3.scaleBand()
        .range([mheight, 0])
        .domain(calendar.yaxis)
        .padding(0.2);

//console.log(yscale.tickSize());

        // create staggered display to avoid too many overlaps
        // keep track of offset for each index on the y axis
        offsetWidthInPx = 15;
        stackHeight = 8;
        offsets = {};
        for(i=0; i<calendar.yaxis.length;i++) {
          offsets[calendar.yaxis[i]] = 0;
        }

	d3.select('#chartarea')
        .selectAll('text')
	.data(calendar.data)
	.enter()
	.append('text')
        .attr('x', function(d) {
          arrival = new Date(parseInt(d.Arrival.replace('/Date(', '').replace(')/', '')));
          return xscale(arrival)+15;
        })
        .attr('y', function(d) {

          // start y values 25px off the axis for that team
          yval = yscale(d.Team)+25;

          // add the offset to create staggered display
          offsetkey = d.Team;
          yval += offsetWidthInPx * offsets[offsetkey];

          // increment or reset the staggered y values
          if(offsets[offsetkey]+1 >= stackHeight) {
            offsets[offsetkey] = 0;
          } else {
            offsets[offsetkey]++;
          } 

          return yval;
        })
	.text(function(d) {
		return d.Title;
	})

        // use inline styles so the svg can be right-clicked -> save image as and e-mailed ...
        .style('font-family', '"Segoe UI","Segoe",Tahoma')
        .style('font-size', '10px')
        .style('fill', 'black')
        .style('text-anchor', 'start')
        .append('title')
          .text(function(d) {
            return d.Title; 
        });

        // reset staggered offsets to begin the rectangles
        for(i=0; i<calendar.yaxis.length;i++) {
          offsets[calendar.yaxis[i]] = 0;
        }

      enter = d3.select('#chartarea')
        .selectAll('rect')
        .data(calendar.data)
        .enter()
        .append('rect')
        .attr('x', function(d) {
          arrival = new Date(parseInt(d.Arrival.replace('/Date(', '').replace(')/', '')));
          return xscale(arrival);
        })
        .attr('y', function(d) {
          yval = yscale(d.Team)+17;

          offsetkey = d.Team;
          yval += offsetWidthInPx * offsets[offsetkey];

          if(offsets[offsetkey]+1 >= stackHeight) {
            offsets[offsetkey] = 0;
          } else {
            offsets[offsetkey]++;
          } 

          return yval;
        })
        .attr('height', '10')
        .attr('width', '10')
        .style('fill', function(d) {
	  if(d.ParentId == 0) {
            return '#000';
          } else {
            return '#ccc';
          }
        })
        .style('stroke', 'black')
        .style('stroke-width', '1px');

        enter.append('title')
        .text(function(d) {
          return spEpochToLocale(d.Arrival) + (d.Dependent ? ', offset: ' + d.OffsetDays + ' days -> ' + calendar.titles['id_' + d.ParentId] : ''); 
        });

        // attach drag events
        enter.call(d3.drag()
          .on('start', dragstarted)
          .on('drag', dragging)
          .on('end', dragended)
        );
        
      // must be last after appending data
      xAxis = chart.append('g').call(d3.axisBottom(xscale));
      xAxis.attr('transform', 'translate(0, ' + mheight + ')');

      // append an x axis for each lane at the zero value for that team
      for(i=0; i<calendar.yaxis.length;i++) {
          axis = chart.append('g').call(d3.axisBottom(xscale));
          axis.attr('transform', 'translate(0, ' + yscale( calendar.yaxis[i] ) + ')');
      }

      // append y axis
      chart.append('g').call(d3.axisLeft(yscale));

      // add a title to the diagram
      d3.select('#chartarea')
	.append('text')
        .attr('x', 10)
        .attr('y', -10)
	.text(charttitle)
        .style('fill', '#0072c6')
        .style('font-family', '"Segoe UI","Segoe",Tahoma');

      resolve(calendar);

  });
}

function updateListItem(data) {

  return new Promise(function(resolve, reject) {

    // id goes in the url
    urlstring = 'http://sharepoint/_vti_bin/listdata.svc/Calendar(' + data.Id + ')'
    delete data.Id;

    $.ajax({
      url: urlstring,
      type: 'POST',
      contentType: 'application/json;odata=verbose',
      processData: false,
      data: JSON.stringify(data),
      headers:{
            'Accept':'application/json;odata=verbose',
            'X-RequestDigest':$('#__REQUESTDIGEST').val(),
            'X-HTTP-Method':'MERGE',
            'If-Match':'*'
      },
      success: function(data) {
        // reload the entire thing
        init();
        resolve(data);
      },
      error: function(jqxhr, status, exception) {
        reject(exception);
      }
    });
  });

}

//////////////////// UTILS ////////////////////

function dragstarted(d) {
    xstart = d3.select(this).attr('x');
    d3.select(this).raise().classed('active', true);

    d3.select('#chartarea')
	.append('text')
        .attr('x', function(d) {
          return xstart;
        })
        .attr('y', function(d) {
          return (d3.select(this).attr('y')-5);
        })
	.attr('id', 'offsethelper')
	.text(function(d) {
          return '';
	})
        .style('font-family', '"Segoe UI","Segoe",Tahoma');
}

function dragging(d) {
    d3.select(this).attr('x', d.x = d3.event.x);

    xstop = d3.select(this).attr('x');
    startDate = xscale.invert(xstart);
    stopDate = xscale.invert(xstop);
    difftime = stopDate - startDate;
    diffdays = Math.ceil(difftime / (1000 * 60 * 60 * 24));

    d3.select('#offsethelper')
        .attr('x', function(d) {
          return d3.event.x;
        })
        .attr('y', function(d) {
          return (d3.event.y - 5);
        })
	.attr('id', 'offsethelper')
	.text(function(d) {
          return 'Change offset ' + (diffdays > 0 ? '+' + diffdays : diffdays);
	});

}

function dragended(d) {

    d3.select('#offsethelper').remove();

    xscale = d3.scaleTime()
        .range([0, mwidth])
        .domain(calendar.xaxis);

    xstop = d3.select(this).attr('x');
    startDate = xscale.invert(xstart);
    stopDate = xscale.invert(xstop);
    difftime = stopDate - startDate;
    diffdays = Math.ceil(difftime / (1000 * 60 * 60 * 24));

    // either update offset from parent or date itself
    if(d.Dependent) {
      if(confirm('Change offset ' + (diffdays > 0 ? '+' + diffdays : diffdays))) {
        updateObject = {};
        updateObject.Id = d.Id;
        updateObject.OffsetDays = (parseInt(d.OffsetDays) + diffdays);

        updateListItem(updateObject);
      } else {
        d3.select(this).attr('x', d.x = xstart);
      }
    } else {
      if(confirm('Save new date ' + stopDate.getDate() + '/' + (stopDate.getMonth()+1) + ' ' + stopDate.getFullYear())) {
        updateObject = {};
        updateObject.Id = d.Id;
        updateObject.Arrival = isoString(stopDate);
        updateObject.EndTime = isoString(stopDate.setDate(stopDate.getDate() + 1));

        updateListItem(updateObject);
      } else {
        d3.select(this).attr('x', d.x = xstart);
      }
    }

    d3.select(this).classed('active', false);
}

function pad(n) { 
  return n < 10 ? '0' + n : n; 
}

// handle Sharepoint dateTime format: /Date(1583766564172)/
function spEpochToLocale(spEpoch) {
  epoch = spEpoch.replace('/Date(', '').replace(')/', '');
  date= new Date(parseInt(epoch));

  monthNames = [
    'January', 'February', 'March',
    'April', 'May', 'June', 'July',
    'August', 'September', 'October',
    'November', 'December'
  ];

  day = date.getDate();
  monthIndex = date.getMonth();
  year = date.getFullYear();

  return day + ' ' + monthNames[monthIndex] + ' ' + year;
}


function isoString(pdate) {

    date = new Date(pdate);
    return date.getUTCFullYear() + '-'
        + pad(date.getUTCMonth() + 1) + '-'
        + pad(date.getUTCDate()) + 'T'
        + pad(date.getUTCHours()) + ':'
        + pad(date.getUTCMinutes()) + ':'
        + pad(date.getUTCSeconds()) + 'Z';
}
