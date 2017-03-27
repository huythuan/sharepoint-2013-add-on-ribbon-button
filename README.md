# sharepoint-2013-add-on-ribbon-button
Paste the following code to javascript file of site colelction:

jQuery(document).ready(function($) {

	// Makes page layouts compatible with other master pages that may have a container at a higer level
	checkForChildContainer();
	
	
	//Creating Ribbon Tabs and Buttons in JavaScript
    initEmbedTabRibbon();
    
    // check existing event
   //This variable will hold a reference to the Announcements list items collection  
    var returnedItems = null;  
    //alert();
   // checkEvent();
});
function PreSaveAction(){
  $startDate = getDateField("Start Date");
  $endDate = getDateField("End Date");
  $conferenceRoom = getInputField("Conference Room");
  
  
  queryListItems();
  return false;
}
/**
 * Get time from a SharePoint Date / time field 
 * @field: title of input field
 * return @string of unix time stamp of input date field
 */
function getDateField(field){
  var ddate = $(":input[title='" + field + "']").val();
  var ddateArray = ddate.split('\/');
  // date looks like DD/MM/YYYY
  var dday = ddateArray[1];
  var dmonth = ddateArray[0];
  var dyear= ddateArray[2];
  var dDateID = $(":input[title='"+field+"']").attr("id");
  var dDateHours = $(":input[id='"+ dDateID + "Hours" +"']").val();
  //delete the “:”
  dDateHours = dDateHours.substring(0,2)
  var dDateMinutes = $(":input[id='"+ dDateID + "Minutes" +"']").val();
  var departuredate = new Date(dyear, dmonth - 1, dday, dDateHours, dDateMinutes,0).getTime()/1000;
  //return departuredate;
  return departuredate;
}

/**
 * get input field value
 * @field: title of input field
 * return string @data of input field
 */
function getInputField(field) {
  var data = $(":input[title='" + field + "']").val();
  return data;
}

//This function fires when the query fails  
function onFailedCallback(sender, args) {  
 //Formulate HTML to display details of the error  
 var markup = '<p>The request failed: <br>';  
     markup += 'Message: ' + args.get_message() + '<br>';  
 //Display the details  
     displayDiv.innerHTML = markup;  
}  

//This function fires when the query completes successfully  
function onSucceededCallback(sender, args) {  
         //Get an enumerator for the items in the list  
         var enumerator = returnedItems.getEnumerator();  
         //Formulate HTML from the list items  
         var markup = 'Items in the Announcements list that start with "T": <br><br>';  
         //Loop through all the items  
         while (enumerator.moveNext()) {  
         var listItem = enumerator.get_current();  
         var unixStartTime = (new Date(listItem.get_item('StartDate')).getTime()/1000);
         var unixEndTime = (new Date(listItem.get_item('_EndDate')).getTime()/1000);
         var roomConference =  listItem.get_item('Conference_x0020_Room');

        alert(roomConference);
      }  
}  

//This function loads the list and runs the query asynchronously  
function queryListItems() {  

         //Get the current context  
         var context = new SP.ClientContext();  
         //Get the Announcements list. Alter this code to match the name of your list  
         var list = context.get_web().get_lists().getByTitle('Caltesting');  
         //Create a new CAML query  
         var caml = new SP.CamlQuery();  
         //Create the CAML that will return only items with the titles that begin with 'T'  
        
        var queryString = "<View><Query><Where>" +
                          "<Geq><FieldRef Name='Conference_x0020_Room'/><Value Type='Text'>" + 'B-149' + "</Value></Geq>" + 
                          //"<Gt><FieldRef Name='_EndDate'/><Value Type='DateTime'><Today/></Value></Gt>" + 
                          //"<Geq><FieldRef Name='ID'/><Value Type='Number'>1</Value></Geq>" +
                          "</Where></Query></View>";
        
        var queryString1 = '<View><Query><Where><Geq>' + 
                           '<FieldRef Name=\'ID\'/><Value Type=\'Number\'>1</Value>' +
                           '</Geq></Where></Query></View>'
        
         caml.set_viewXml(queryString );  
        
        //Specify the query and load the list oject  
         returnedItems = list.getItems(caml);  
         context.load(returnedItems);  
         //Run the query asynchronously, passing the functions to call when a response arrives  
         context.executeQueryAsync(onSucceededCallback, onFailedCallback);  
}  

    























/**
 * Init ribbon for creating tab
 */
function initEmbedTabRibbon() {
  SP.SOD.executeOrDelayUntilScriptLoaded(function () {

  var pm = SP.Ribbon.PageManager.get_instance();

  pm.add_ribbonInited(function () {
    createEmbedTab();
  });

  var ribbon = null;
  try {
    ribbon = pm.get_ribbon();
  }
  catch (e) { }

    if (!ribbon) {
      if (typeof (_ribbonStartInit) == "function")
        _ribbonStartInit(_ribbon.initialTabId, false, null);
      }
    else {
        createEmbedTab();
      }
  }, "sp.ribbon.js");
}

/**
 * Creating Ribbon Tabs and Buttons in JavaScript
 */
function createEmbedTab() {
    var ribbon = SP.Ribbon.PageManager.get_instance().get_ribbon();
    if (ribbon !== null) {
        var ribbonTab = ribbon.getChild("Ribbon.EditingTools.CPInsert"); // CUI.Tab
        
        //Add youtube embed group
        var groupYoutube = new CUI.Group(ribbon, 'Ribbon.EditingTools.CPInsert.Youtube', 'Embed', 'Use this group for embed operations', 'Youtube.Group.Command', null);
     
        var layout = new CUI.Layout(ribbon, 'Youtube.Layout', 'Youtube.Layout');
        groupYoutube.addChild(layout);
        var section = new CUI.Section(ribbon, 'Youtube.Section', 2, 'Top'); //2==OneRow
        layout.addChild(section);
        groupYoutube.selectLayout(layout.get_title(), layout);
        
        var controlProperties = new CUI.ControlProperties();
        var button = new CUI.Controls.Button(ribbon, 'Youtube.Button', controlProperties);
        var controlComponent = new CUI.ControlComponent(ribbon, 'Youtube.ControlComponent', 'Large', button);
        var row1 = section.getRow(1);
        row1.addChild(controlComponent);
        ribbonTab.addChild(groupYoutube);
        
        //add youtube icon and link to button
        var htmlYoutube = "";
        htmlYoutube += "<a href='#' onclick='addYoutubeEmbedCode(); return false;'>";
        htmlYoutube += "<img src='/_catalogs/masterpage/images/tliyoutube2.png' /><br>Youtube</a>";
        var buttonElement = button.getDOMElementForDisplayMode("Large");
        $(buttonElement).html(htmlYoutube);
        
        //Add google map embed group
        var ribbonTabMap = ribbon.getChild("Ribbon.EditingTools.CPInsert"); // CUI.Tab
        var groupGoogleMap = new CUI.Group(ribbon, 'Ribbon_EditingTools_CPInsert_Google_Map', 'Embed', 'Use this group for embed operations', 'Map.Group.Command', null);
        
        var layoutGoogleMap = new CUI.Layout(ribbon, 'Map.Layout', 'Map.Layout');
        groupGoogleMap.addChild(layoutGoogleMap);
        var sectionGoogleMap = new CUI.Section(ribbon, 'Map.Section', 2, 'Top'); //2==OneRow
        layoutGoogleMap.addChild(sectionGoogleMap);
        groupGoogleMap.selectLayout(layoutGoogleMap.get_title(), layoutGoogleMap);
        
        var controlGoogleMapProperties = new CUI.ControlProperties();
        var buttonGoogleMap = new CUI.Controls.Button(ribbon, 'GoogleMap.Button', controlProperties);
        var controlGoogleMapComponent = new CUI.ControlComponent(ribbon, 'GoogleMap.ControlComponent', 'Large', buttonGoogleMap);
        var rowGoogleMap = sectionGoogleMap.getRow(1);
        rowGoogleMap.addChild(controlGoogleMapComponent);
        ribbonTabMap.addChild(groupGoogleMap);
        
        //add google map icon and link to button
        var htmlYoutube = "";
        htmlYoutube += "<a href='#' onclick='addGoogleMapEmbedCode(); return false;'>";
        htmlYoutube += "<img src='/_catalogs/masterpage/images/map.png' /><br>Map</a>";
        var buttonMapElement = buttonGoogleMap.getDOMElementForDisplayMode("Large");
        $(buttonMapElement).html(htmlYoutube);

        //Add audio embed group
        var ribbonTabAudio = ribbon.getChild("Ribbon.EditingTools.CPInsert"); // CUI.Tab
        var groupAudio = new CUI.Group(ribbon, 'Ribbon_EditingTools_CPInsert_Audio', 'Embed', 'Use this group for embed operations', 'Audio.Group.Command', null);
        
        var layoutAudio = new CUI.Layout(ribbon, 'Audio.Layout', 'Audio.Layout');
        groupAudio.addChild(layoutAudio);
        var sectionAudio = new CUI.Section(ribbon, 'Audio.Section', 2, 'Top'); //2==OneRow
        layoutAudio.addChild(sectionAudio);
        groupAudio.selectLayout(layoutAudio.get_title(), layoutAudio);
        
        var controlAudioProperties = new CUI.ControlProperties();
        var buttonAudio = new CUI.Controls.Button(ribbon, 'Audio.Button', controlProperties);
        var controlAudioComponent = new CUI.ControlComponent(ribbon, 'Audio.ControlComponent', 'Large', buttonAudio);
        var rowAudio = sectionAudio.getRow(1);
        rowAudio.addChild(controlAudioComponent);

        //add google map icon and link to button
        var htmlAudio = "";
        htmlAudio += "<a href='#' onclick='addAudioEmbedCode(); return false;'>";
        htmlAudio += "<img src='/_catalogs/masterpage/images/audio_icon.png' /><br>Audio</a>";
        var buttonAudioElement = buttonAudio.getDOMElementForDisplayMode("Large");
        $(buttonAudioElement).html(htmlAudio);
        ribbonTabAudio.addChild(groupAudio);

    }
}

/**
 * Add Audio embed code to content
 */
function addAudioEmbedCode() {
  var audioLink = prompt("Enter an audio link", "");
  if (audioLink !=null && audioLink!=""){
    var content = '';          
    content += '<p>';
    content += '<audio controls="controls"><source src="' + audioLink + '"' + ' type="audio/mpeg">';
    content += 'Click <a href="' + audioLink + '"> here </a> to listen';
    content += '</audio>';
    content += '</p>';
    $('#ms-rterangecursor-start').parent().html(content);
  }  
}

/**
 * Add Google Map embeded code to content
 */
function addGoogleMapEmbedCode() {
  var url = prompt("Enter an address", "");
  var width = prompt("Enter width", "100%");
  var height = prompt("Enter height", "450");
  if (url !=null && url !="" && width !=null && width !="" && height !=null && height !=""){
    url = url.trim();
    var content = '<p><iframe width="' + width  + '" height="' + height + '"';
      content += 'src="//maps.google.com/maps?q=' + url + '&num=1&t=m&ie=UTF8&z=14&output=embed"'
      content += 'frameborder="0" scrolling="no" style="border:0"></iframe></p>';  
    $('#ms-rterangecursor-start').parent().html(content); 
  }  
}
/**
 * Add youtube embed code to content
 */
function addYoutubeEmbedCode() {
  var youtubeUrl = prompt("Youtube URL", "");
  if (youtubeUrl != null) {
    var url = youtubeUrl.trim();
    var regExURL=/v=([^&$]+)/i;
    var id_video=url.match(regExURL);
	if(id_video==null || id_video=='' || id_video[0]=='' || id_video[1]==''){
      alert("Invalid youtube url");
	  return false;
	}
    var width =  '560';
    var height = '315';
    var content = '';
    content += '<p class="pm-video">';
    content += '<iframe class="youtube-field-player" width="' + width + '" height="'
               + height + '" src="//www.youtube.com/embed/' + id_video[1] 
               + '?rel=0&modestbranding=1&theme=light&color=white&wmode=opaque&autoplay=0&showinfo=0"'
               + 'frameborder="0" style="width: 100%;" allowfullscreen="1">';
    content += '</iframe>';
    content += '</p>';
    $('#ms-rterangecursor-start').parent().html(content);   
  }
}

function checkForChildContainer() {
	$('.container').each(function(){	
		$(this)
		.find('.container')
		.removeClass('container')
		.addClass('newGroup');
	
	});
}
