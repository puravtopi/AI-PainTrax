// var _cal;
// var _cal_tab;
// var _cal_body;
var _eventData;
var _DEFAULT_TIME = 30;
var calChange = {};
var CalAdd = {};
var SelectDate = null;
function padLeft(str, char, len) {

    var pad = "";
    for (i = 1; i <= len; i++) {
        pad += char;
    }
    // alert(pad);
    var ans = pad.substring(0, pad.length - str.length) + str;
    return ans;
}
function showCalendar(div,divMonth,divYear)
{
    _cal_Month = document.getElementById(divMonth);
    _cal_Year = document.getElementById(divYear);

    _cal = document.getElementById(div);

    _cal = document.getElementById(div);
    _cal.innerHTML = "";
    _cal_tab = document.createElement("table");
    _cal_tab.id = "_cal_table";    
    _cal_tab.className = "table table-bordered mytable";
    _cal_body = document.createElement("tbody");
    _cal_body.id = "_cal_body";
    _cal.appendChild(_cal_tab);    
    _cal_tab.appendChild(_cal_body);

    const caption = document.createElement("caption");
    caption.id="_cal_caption";
    _cal_tab.appendChild(caption);

    const selectMonth= document.createElement("select");
    const selectYear= document.createElement("select");
    const labelMonth= document.createElement("label");
    const labelYear= document.createElement("label");

    selectMonth.className = "form-select";
    selectYear.className = "form-select";

    labelMonth.innerHTML=" Month : ";
    labelYear.innerHTML=" Year : ";
    //caption.appendChild(labelMonth);
    //caption.appendChild(selectMonth);
    //caption.appendChild(labelYear);
    //caption.appendChild(selectYear);
    _cal_Month.appendChild(selectMonth);
    _cal_Year.appendChild(selectYear);
    selectMonth.onchange = onSelectChange;
    selectYear.onchange = onSelectChange;
    //selectMonth.onchange =function() { onSelectChange(_cal,_cal_tab, _cal_body);};
    //selectYear.onchange = function() { onSelectChange(_cal,_cal_tab, _cal_body);};

    selectMonth.id="_select_month";
    selectYear.id="_select_year";
    
    const months = new Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December");
    for (let l = 0; l < months.length; l++) 
    {
        var option = document.createElement("option");
        option.value=l+1;
        txt = document.createTextNode(months[l]);
        option.appendChild(txt);
        selectMonth.insertBefore(option, selectMonth.lastchild);
    }
    //const year = new Array("2000","2001","2002","2003","2004","2005","2006","2007","2008","2009","2010","2011","2012","2013","2014","2015","2016","2017","2018","2019","2020","2021","2022","2023","2024","2025","2026","2027","2028","2029","2030","2031","2032","2033","2034","2035","2036","2037","2038","2039","2040","2041","2042");
    for (let y = 0; y < 10; y++) 
    {
        var option1 = document.createElement("option");
        option1.value=y+2020;
        txt1 = document.createTextNode(y+2020);
        option1.appendChild(txt1);
        selectYear.insertBefore(option1, selectYear.lastchild);
    }
    var curDate=new Date();
    selectMonth.value=curDate.getMonth()+1;
    selectYear.value=curDate.getFullYear();
    
    //myFunction();
    onSelectChange();
}
//function myFunction() 
function onSelectChange()
{
    
   // const _cal = document.getElementById(div);
    //const _cal_tab = document.getElementById("_cal_table");
    const _cal_body = document.getElementById("_cal_body");

    //const selectMonth= document.getElementById("_select_month");
    //const selectYear=  document.getElementById("_select_year");
    const selectMonth= document.getElementById("_select_month");
    const selectYear=  document.getElementById("_select_year");

    var mo=selectMonth.value;
    var ye=selectYear.value;
    //var date = new Date();
    var fd = new Date(ye, mo-1,1).getDay();
    var ld = new Date(ye, mo,0).getDate();
    var pld = new Date(ye, mo-1,0).getDate();
    _cal_body.innerHTML="";    
    const days = ["Sun", "Mon", "Tue", "Wed", "Thr", "Fri", "Sat"];
    let k = 1;
    let l=1;
    outer:
    var row = document.createElement("tr");
    var data = document.createElement("th");
    const months = new Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December");
    data.colSpan="7";   
    data.classList.add("_cal_head");            
    data.innerHTML = months[mo-1] + " "+ye;
    row.appendChild(data);
    _cal_body.appendChild(row);
    for (let i = 0; i <=6 ; i++) 
    {
        const row = document.createElement("tr");
        _cal_body.appendChild(row);
        if (i == 0) 
        {
            for (let j = 1; j <= 7; j++) 
            {
                const data = document.createElement("th");
                data.classList.add("_cal_th");        
                data.innerHTML = days[j - 1];
                row.appendChild(data);
            }
        }
        else 
        {    
            for (m = 1; m <= fd; m++) 
            {   
                mon=parseInt(mo)-1;
                var year=ye;
                if(mon<1){
                    year--;
                    mon=12;
                }
                if(i!=1)
                {
                    fd=0;
                    break;
                }
                addData(row,((pld-fd)+m)+"",mon+"",year,"_cal_prevMonth");
            }
            for (let m = 1; m <= 7-fd; m++,k++) 
            {
                if (k > ld)
                {
                    mon=parseInt(mo)+1;
                    var year=ye;
                    if(mon>12){
                        year++;
                        mon=1;
                    }                        
                    addData(row,l,mon+"",year,"_cal_nextMonth");
                    l++;    
                }else{
                    addData(row,k,mo,ye);
                }
            }
        }
      //  console.log("1");
      // console.log(_eventData);
       // setData(_eventData,false);
    }
  //  setData(_eventData,false);
    
    function addData(row,day,month,year,dataClass="")
    {
        const today= new Date();
        const curDate= today.getFullYear()+padLeft((today.getMonth()+1)+"","0",2)+padLeft(today.getDate()+"","0",2);
        const tabDate=year+padLeft(month,"0",2)+padLeft(day+"","0",2)
        const data = document.createElement("td");
        const label = document.createElement("label");
        const inputhidden = document.createElement("input");
        if (dataClass) {
            data.classList.add("_cal_nextMonth");        
        }
        if(curDate===tabDate )   {
            data.classList.add("_cal_curDate");
        }
        data.id=tabDate;
        inputhidden.type="hidden";
        inputhidden.value=tabDate;        
        label.innerHTML = day;        
        label.classList.add("dayLabel");
        data.appendChild(inputhidden);
        data.appendChild(label);
        row.appendChild(data);
        
        if (_eventData) {
           // debugger
            const findData = _eventData.find(obj => obj.date == tabDate);
            if (findData) {
                const label = document.createElement("label");
                label.id = "label" + findData.date;
                label.classList.add("badge");
                label.classList.add("bg-primary");
                label.classList.add("noOfEvent");
                label.innerText = findData.appointments;
                data.insertBefore(label, data.firstChild);
                //data.appendChild(label);
            }
        }
        data.onclick = function () { dateClicked(tabDate); };
    }  
}
function setData(data)
{
    _eventData = data;
    
    onSelectChange();
}
function setAppData(data) {
    console.log(data);
    data.forEach((appoint) => {
        console.log(appoint);
       // console.log(appoint);

        const datatd = document.getElementById(appoint.appointmentStart);
        const statustd = document.getElementById('status_' + appoint.appointmentStart);
        const removetd = document.getElementById('remove_' + appoint.appointmentStart);
        if (datatd) {
            console.log("datafound")
            const btnRemove = document.createElement("input");
            btnRemove.type = "button"
            btnRemove.value = "X";
            btnRemove.classList.add("btn");
            btnRemove.classList.add("btn-sm");
            btnRemove.classList.add("btn-danger");
            btnRemove.classList.add("appBtn");
            btnRemove.onclick = function () { calRemove(appoint); }
            removetd.appendChild(btnRemove);
            const appLabel = document.createElement("div");
            //alert(appoint.AppointmentNote);
            appLabel.innerText = appoint.fname +' '+ appoint.lname +(( appoint.appointmentNote.length>0)? " | " : "") + appoint.appointmentNote +  " | " + appoint.ProviderName;
            appLabel.classList.add("appLabel");
            console.log(appoint);
            datatd.appendChild(appLabel);
            const btnStatus = document.createElement("input");
            btnStatus.type = "button"
            btnStatus.value = appoint.status;
            btnStatus.classList.add("btn");
            btnStatus.classList.add("btn-sm");
            btnStatus.classList.add("btn-primary");
            btnStatus.classList.add("appBtn");
            //btnStatus.classList.add("appBtn");
            btnStatus.onclick = function () { calChangeStatus(appoint); }
            statustd.appendChild(btnStatus);

            //document.getElementById("btn_" + appoint.AppointmentStart).remove();
            appLabel.onclick = function () { calChange(appoint); }
 

        }
    });
    
}

function showApp(selectedDate, div) {
    var start = 480, end = 480;
    _app = document.getElementById(div);
    _app.innerHTML = "";
    const caption = document.createElement("div");
    caption.id = "_app_caption";
    _app.appendChild(caption);
    const anchor = document.createElement("a");
    anchor.href = "#";
    anchor.onclick = function () { calClicked(); };
    anchor.classList.add("btn");
    anchor.classList.add("btn-primary");
    anchor.innerText = "Show Calendar";
    caption.appendChild(anchor);



    
    _app_tab = document.createElement("table");
    _app_tab.id = "_app_table";
    _app_tab.className ="table table-bordered daytable"
    _app_head = document.createElement("thead");
    _app_head.id = "_app_head";
    _app_tab.appendChild(_app_head);
    _app_body = document.createElement("tbody");
    _app_body.id = "_app_body";
    _app.appendChild(_app_tab);
    _app_tab.appendChild(_app_body);
    SelectDate = selectedDate;
    const headRow = document.createElement("tr");
    const headDate = document.createElement("th");
    headDate.colSpan = 2;
    headDate.classList.add("_app_th");
    if (selectedDate) {
        headDate.innerHTML = selectedDate.substring(4, 6) + " / " + selectedDate.substring(6, 8) +" / "+ selectedDate.substring(0, 4);
    }
    
    headRow.appendChild(headDate);
    _app_head.appendChild(headRow);


    const headRow2 = document.createElement("tr");
    const headTime = document.createElement("th");
    const headAdd = document.createElement("th");
    const headName = document.createElement("th");
    const headStatus = document.createElement("th");
    const headRemove = document.createElement("th");
    headTime.innerHTML = "Time";
    headAdd.innerHTML = "New Appointment";
    headName.innerHTML = "Patient Name";
    headStatus.innerHTML = "Status";
    headRemove.innerHTML = "Remove";
    headRow2.appendChild(headTime);
    headRow2.appendChild(headAdd);
    headRow2.appendChild(headName);
    headRow2.appendChild(headStatus);
    headRow2.appendChild(headRemove);
    _app_head.appendChild(headRow2);

   // console.log("test");
   /* for (var i = 720; i <= 750; i += _DEFAULT_TIME) {
        getRow(i, " AM",i-720);
    }*/
    for (var i = daystart; i <= 690; i += _DEFAULT_TIME) {
        getRow(i," AM",i);
    }
    for (var i = 720; i <= 750; i += _DEFAULT_TIME) {
        getRow(i, " PM",i);
    }
    for (var i = 60; i <= dayend; i += _DEFAULT_TIME) {
        getRow(i, " PM",i+720);
    }

    function getRow(val,noon,val24) {
        const row = document.createElement("tr");
        const time = document.createElement("td");
        time.classList.add("_app_td_time");
        time.classList.add("bg-primary");
        time.innerHTML = minsToHm(val) + noon ;
        row.appendChild(time);
        const add = document.createElement("td");
        const data = document.createElement("td");
        const status = document.createElement("td");
        const remove = document.createElement("td");
        data.classList.add("_app_td_data");
        remove.classList.add("_app_td_remove");
        data.id = padLeft(minsToHm(val24) + "", "0", 5);
        status.id = "status_" + padLeft(minsToHm(val24) + "", "0", 5);
        remove.id = "remove_" + padLeft(minsToHm(val24) + "", "0", 5);

        const btnNew = document.createElement("input");
        btnNew.id = "btn_"+padLeft(minsToHm(val24) + "", "0", 5);
        btnNew.type = "button"
        btnNew.value = "Add";
        btnNew.classList.add("btn");

        btnNew.classList.add("btn-sm");
        btnNew.classList.add("btn-primary");
        btnNew.classList.add("appBtn");
        
        btnNew.onclick = function () { calAdd(minsToHm(val) + noon, SelectDate); }
        
        add.classList.add("_app_td_add");
        status.classList.add("_app_td_status");
        add.appendChild(btnNew);
        row.appendChild(add);
        row.appendChild(data);
        row.appendChild(status);
        row.appendChild(remove);
        _app_body.appendChild(row);
    }
    

}
function minsToHm(d) {
    d = Number(d);
    var h = Math.floor(d / 60);
    var m = Math.floor(d % 60);
    return parseFloat(h + "." + m).toFixed(2).replace(".", ":");
}
function setZero(time) {

    return (time.length < 5 ? "0" + time : time);
}

function getTimeSet() {
    optionstr = "";
    for (var i = daystart; i <= 690; i += _DEFAULT_TIME) {
        
        
        optionstr += "<option value='" + setZero(minsToHm(i)) + "'>" + setZero(minsToHm(i)) + " AM" + "</option>";
    }
    for (var i = 720; i <= 750; i += _DEFAULT_TIME) {
        optionstr += "<option value='" + setZero(minsToHm(i)) + "'>" + setZero(minsToHm(i)) + " PM" + "</option>";
    }
    for (var i = 60; i <= dayend; i += _DEFAULT_TIME) {
        optionstr += "<option value='" + minsToHm(i + 720) + "'>" + setZero(minsToHm(i)) + " PM" + "</option>";
    }
    return optionstr;
}