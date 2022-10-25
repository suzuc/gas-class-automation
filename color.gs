function decideColor(location, jugyoukoumoku) {  
  
  const gAccount = "";
  var calendar = CalendarApp.getCalendarById(gAccount);

  if (jugyoukoumoku == "試験"){
    return "RED"

  }else if(jugyoukoumoku == "TBL"){
    return "ORANGE"

  }else if(~jugyoukoumoku.indexOf("実習")){
    return "ORANGE"

  }else if(location == "オンライン"){
    return "GREEN"

  }else{
    return "PALE_RED"
  };
}
