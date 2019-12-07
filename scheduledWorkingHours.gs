function scheduledWorkingHours(){
  const sheet = SpreadsheetApp.getActive().getSheetByName("PlayGround");
  setHeader(sheet)

  var startDate = Moment.moment("2019-12-01 00:00:00")
  var monthCount=13
  var holidaysList=[]
  const calendar = CalendarApp.getCalendarsByName('日本の祝日')[0];

  var count_cloneDate=startDate.clone()
  holidaysList=holidayList(count_cloneDate,calendar)
  sheet.getRange(2,8,holidaysList.length,4).setValues(holidaysList)


  for (var i = 0, len = monthCount; i < len; i++) {
    var dateCount=startDate
    var param_daysInMonth=dateCount.daysInMonth()
    var param_weekendInMonth=weekendInMonth(dateCount)
    var param_validHolidayInMonth=validHolidaysInMonth(dateCount,calendar)
    var param_workingDay=param_daysInMonth-(param_weekendInMonth+param_validHolidayInMonth)
    var param_workingTime=param_workingDay*8
    sheet.getRange(2+i,1).setValue(dateCount.format("YYYYMM"))
    sheet.getRange(2+i,2).setValue(param_daysInMonth)
    sheet.getRange(2+i,3).setValue(param_weekendInMonth)
    sheet.getRange(2+i,4).setValue(param_validHolidayInMonth)
    sheet.getRange(2+i,5).setValue(param_workingDay)
    sheet.getRange(2+i,6).setValue(param_workingTime)
    dateCount.add(1,"M").format("YYYYMM")
  }
}

function setHeader(sheet){
  var list=[["対象月","総日数","土日の数","休みが増える祝日数","所定労働日数","所定労働時間","","祝日名称","日付","休みが増える","曜日"]]
  sheet.getRange(1,1,1,list[0].length).setValues(list)
}

function weekendInMonth(dateCount){
  var dateCount_clone=dateCount.clone()
  var holidays=0
  for (var i = 0, len = dateCount_clone.daysInMonth(); i < len; i++) {
    if(dateCount_clone.day()==0 || dateCount_clone.day()==6){
      Logger.log(dateCount_clone.format("YYYY-MM-DD"))
      holidays+=1
    }
   dateCount_clone.add(1,"Days")
  }
  return holidays
  //var countDayOfWeek=dateCount.daysInMonth()+dateCount.day()
  //var saturday=Math.floor(countDayOfWeek/6)
  //var sunday=Math.floor(countDayOfWeek/7)
  //return saturday+sunday
}


function validHolidaysInMonth(dateCount,calendar){
  var start_cloneDate=dateCount.clone()
  var end_cloneDate=dateCount.clone()
  var start=start_cloneDate.toDate()
  var end  =end_cloneDate.add(1,"Months").toDate()

  var events = calendar.getEvents(start,end);
  var holidayCount=0
  for (var i = 0, len = events.length; i < len; i++) {
    if(events[i].getStartTime().getDay()==0 || events[i].getStartTime().getDay()==6){
    }else{
      holidayCount++
    }
  }
  return holidayCount
}

function holidayList(dateCount,calendar) {
  var holidays=[]

  var start_cloneDate=dateCount.clone()
  var end_cloneDate=dateCount.clone()
  var start=start_cloneDate.toDate()
  var end  =end_cloneDate.add(2,"Years").toDate()

  var events = calendar.getEvents(start,end);
  for (var i = 0, len = events.length; i < len; i++) {
    var yyyy=events[i].getStartTime().getFullYear().toString()
    var mm  =events[i].getStartTime().getMonth()+1
    mm=mm.toString()
    if(mm.length===1){
      mm="0"+mm
    }
    var dd  =events[i].getStartTime().getDate().toString()
    if(dd.length===1){
      dd="0"+dd
    }
    var yyyymmdd=yyyy+mm+dd

    var dayOfWeek=dict_NumberToDayOfWeek()[events[i].getStartTime().getDay()]

    var validholiday="○"
    if(events[i].getStartTime().getDay()==0 || events[i].getStartTime().getDay()==6){validholiday="×"}
    holidayParam=[events[i].getTitle(),yyyymmdd,validholiday,dayOfWeek]
    holidays.push(holidayParam)
  }
  return holidays
}
