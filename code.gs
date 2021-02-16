function get_column_index(active_sheet,column_name) 
{
  var column_arr=[];
  //var headers = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  var headers=active_sheet.getRange(1, 1, 1, 50);
  values = headers.getValues();
  for (var j in values[0]) 
  {
    column_arr.push(values[0][j])  
  }
  var column_index = column_arr.indexOf(column_name);
  return column_index+1;
}



function time_stamp(active_sheet,row,pst,taget_col_value)
{
    var space="---";
    var time_drafted_col=get_column_index(active_sheet,"Time Drafted") ;
    var existing_value=active_sheet.getRange(row,time_drafted_col).getValue();
    var new_value=taget_col_value+space+pst;
    var new_value_pst=new_value+existing_value
    active_sheet.getRange(row,time_drafted_col).setValue("\n" + new_value_pst);
    return new_value;
}

function store_sheet_data(value,col_name,sheet_name,taget_col_value)
{
  if(taget_col_value !== "Approved (Don’t send)"){
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheet_name); 
  var columns = "ABCDEFGHIJKLMNOPYRSTUVWXYZ"; 
  var target_col_no=get_column_index(sheet,col_name); 
  var target_col = columns.charAt(target_col_no-1);
  var target_range=target_col + "1:" +target_col;
  var target_col_val = sheet.getRange(target_range).getValues();
  var target_col_lr = target_col_val.filter(String).length;
  var target_col_lr_val = sheet.getRange(target_col_lr+1,target_col_no);
  target_col_lr_val.setValue(value);  
}
}
function draft_status(taget_col_value,row,sheet_name,active_sheet,darft_info,pst)
{
  console.log(taget_col_value);
  if(taget_col_value === "Drafted" || taget_col_value === "Sent")
  {
    if(taget_col_value === "Approved (Don’t send)")
    {
      var sent_time_col=get_column_index(active_sheet,"sent_time") ;
      var sent_time=active_sheet.getRange(row,sent_time_col).getValue();
      if(sent_time === '')
      {
        active_sheet.getRange(row,sent_time_col).setValue(pst);
        store_sheet_data(darft_info,"Draft Info","agent_draft_status");
      }
    }

    else if(taget_col_value === "Drafted")
    {
      var drafted_time_col=get_column_index(active_sheet,"drafted_time") ;
      var drafted_time=active_sheet.getRange(row,drafted_time_col).getValue();
      if(drafted_time === '')
      {
        active_sheet.getRange(row,drafted_time_col).setValue(pst);
        store_sheet_data(darft_info,"Draft Info","agent_draft_status");
      }
    }else if(taget_col_value === "Sent")
    {
      var sent_time_col=get_column_index(active_sheet,"sent_time") ;
      var sent_time=active_sheet.getRange(row,sent_time_col).getValue();
      if(sent_time === '')
      {
        active_sheet.getRange(row,sent_time_col).setValue(pst);
        store_sheet_data(darft_info,"Draft Info","agent_draft_status");
      }
    }
  }
  else
  {
    store_sheet_data(darft_info,"Draft Info","agent_draft_status");  
  }
}

function sub_sheet(active_sheet,row,sheet_name,new_value,taget_col_value,pst)
{ 
  var t_header="   ticket_id---";
  var ticket_id_col=get_column_index(active_sheet,"Ticket Id");
  var ticket_id=active_sheet.getRange(row,ticket_id_col).getValue();  
  var drafted_by_col=get_column_index(active_sheet,"Drafted by");
  var drafted_by=active_sheet.getRange(row,drafted_by_col).getValue();  
  var new_value_data=new_value+t_header+ticket_id;
  var start_time_col=get_column_index(active_sheet,"start_time") ;
  var start_time=active_sheet.getRange(row,start_time_col).getValue();
  var new_value_data=new_value+t_header+ticket_id;
  if(start_time === '')
  {
    active_sheet.getRange(row,start_time_col).setValue(pst);
    start_time=active_sheet.getRange(row,start_time_col).getValue();
  }
  if(ticket_id && drafted_by && taget_col_value  && pst && start_time)
  {
    var darft_info=start_time+" --- "+drafted_by+" --- "+ticket_id+" --- "+taget_col_value+" --- "+pst;
    if(taget_col_value !== "Approved (Don’t send)"){
    draft_status(taget_col_value,row,sheet_name,active_sheet,darft_info,pst);
    store_sheet_data(new_value_data,drafted_by,"Work Log");
    }
  } 
}

function agent_mistake(active_sheet,row,col_value)
{
  var draft_edited_col=get_column_index(active_sheet,"Number of times draft was edited") ;
  var value=active_sheet.getRange(row,draft_edited_col).getValue();
  if(col_value !== "Draft Again")
  {
    active_sheet.getRange(row,draft_edited_col).clearContent();
    active_sheet.getRange(row,draft_edited_col).setValue(value+1);
  }   
}


function sheet_info(e,sheet_name)
{
  var active_sheet=e.source.getActiveSheet();
  var active_sheet_name=active_sheet.getName();
  var target_col=get_column_index(active_sheet,"Status");
  var target_col_2=get_column_index(active_sheet,"Escalation Draft Status") ;
  var start_row=2;
  var row=e.range.getRow();
  var col=e.range.getColumn();
  var date = new Date();
  var pst = date.toUTCString();
  var taget_col_value=active_sheet.getRange(row,target_col).getValue();
  var taget_col_2_value=active_sheet.getRange(row,target_col_2).getValue();
  var cell = active_sheet.getActiveCell();
  var cellCol = cell.getColumn();
  if(cellCol == 2 || cellCol == 3){
    var ticket_id_col=get_column_index(active_sheet,"Ticket Id");
    var ticket_id=active_sheet.getRange(row,get_column_index(active_sheet,"Ticket Id")).getValue();  
    var drafted_by_col=get_column_index(active_sheet,"Drafted by");
    var drafted_by=active_sheet.getRange(row,drafted_by_col).getValue(); 
    var drafted_time_col=get_column_index(active_sheet,"Time Drafted");
    var drafted_time_by=active_sheet.getRange(row,drafted_time_col).getValue(); 
    if(ticket_id !="" && drafted_by !="" && taget_col_value !="" && drafted_time_by != ""){
      var split_val = drafted_time_by.split('GMT');
      var check_status_arr = [];
      for(var i=split_val.length;i>=0;i--){
        if(split_val[i]){
          var value_split_val = split_val[i]+'GMT';
          var new_value_data=value_split_val+ '  ticket_id---'+ticket_id;
         
          var taget_col_value = split_val[i].split('---');
          taget_col_value = taget_col_value[0].trim(); 
          store_sheet_data(new_value_data,drafted_by,"Work Log", taget_col_value); 
          var start_time_col=get_column_index(active_sheet,"start_time") ;
          var start_time=active_sheet.getRange(row,start_time_col).getValue();  
          if(taget_col_value === "Drafted" || taget_col_value === "Edited"  || taget_col_value !== "Approved (Don’t send)" || taget_col_value === "Sent")
          {
            value_split_val = split_val[i]+'  GMT'.trim();
            var get_splited_time = value_split_val.split("---");
            if(check_status_arr[get_splited_time[0]] == undefined){
              var darft_info=start_time+" --- "+drafted_by+" --- "+ticket_id+" --- "+taget_col_value+" --- "+get_splited_time[1];
              store_sheet_data(darft_info,"Draft Info","agent_draft_status");
              check_status_arr[get_splited_time[0]] = 1;
            }
          }
        } 
      }
    }
  }
  else if(col === target_col && row >= start_row && active_sheet_name === sheet_name)
  {
    if(taget_col_value === "Drafted" || taget_col_value === "Edited"  || taget_col_value === "Approved (Don’t send)" || taget_col_value === "Sent")
    {
      var new_value=time_stamp(active_sheet,row,pst,taget_col_value);
      sub_sheet(active_sheet,row,sheet_name,new_value,taget_col_value,pst);
    }
    else if(taget_col_value === "Draft Again")
    {
      agent_mistake(active_sheet,row,taget_col_2_value);
    } 
  }
  else if(col === target_col_2 && row >= start_row && active_sheet_name === sheet_name)
  {
    if(taget_col_2_value === "Draft Again")
    {
      agent_mistake(active_sheet,row,taget_col_value);
    } 
  }
}

function onEdit(e) 
{
  var sheet_name="Collective Sheet";
  
  sheet_info(e, sheet_name);
}