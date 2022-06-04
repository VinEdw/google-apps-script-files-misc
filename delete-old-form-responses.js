function onOpen() {
  clear_two_week_old_responses();
}

function clear_two_week_old_responses() {
  let form = FormApp.getActiveForm();
  let responses_arr = form.getResponses();

  let today = new Date(Date.now());
  // today.setDate(today.getDate() + 0); //manual date shifting for testing
  console.log('Today is', today);

  console.log('Deleting responses frome more than 14 days ago...');
  let c = 0
  for (let i = 0; i < responses_arr.length; i++) {
    let response = responses_arr[i];
    let response_id = response.getId();
    let response_date = response.getTimestamp();
    let days_old = (today - response_date)/1000/3600/24;
    console.log(days_old);
    if (days_old > 14) {
      c++;
      console.log('Deleting resposnse', response_id);
      form.deleteResponse(response_id);
    }
  }
  console.log(c, 'responses deleted.')
}