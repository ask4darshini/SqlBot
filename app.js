var builder = require('botbuilder');
var restify = require('restify');
var request = require('request');
var http = require('http');
var pos = require('pos');

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3979, function () {
   console.log('%s listening to %s', server.name, server.url); 
});


// Create chat bot
var connector = new builder.ChatConnector({
   // appId: '60c99b0c-31af-4acf-809b-19d892e7c13f',
   // appPassword: '5dO88PYyQi55oqhb1ncwhT8'
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD	
});
var bot = new builder.UniversalBot(connector);
server.post('/api/messages', connector.listen());


// Create LUIS recognizer that points at our model and add it as the root '/' dialog for our Cortana Bot.
var model = 'https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/1f2f4453-26c5-4782-ab30-58e7a631c9ca?subscription-key=3a93409d819c40abbab8d7130d631f26&staging=true&verbose=true&timezoneOffset=0&spellCheck=false&q=';
var recognizer = new builder.LuisRecognizer(model);
var dialog = new builder.IntentDialog({ recognizers: [recognizer] });
bot.dialog('/', dialog);

/*
bot.use({
	botbuilder: function(session, next) {
	    if (session.message.user.name != "Gagan") {
		session.send("Sorry I am told not to talk to strangers...");
		session.endConversation();
	    } else {
		next();
	    }
	}
});
*/

dialog.onDefault('default', [
	function(session) {
		session.send("Sorry I don't know what you're asking...");
		session.endDialog();
	}
]);

dialog.matches('Greetings', [
	function(session) {
		var name = session.message.user.name.split(" ");
		session.send("Hello %s! My name is Bo. What can I help you with?", name[0]);	
		session.endDialog();
	}
]);

dialog.matches('Access', [
	function(session, args, next) {
		console.log(session.message);
		var msg = new builder.Message(session)
           		.textFormat(builder.TextFormat.xml)
            		.attachments([
           			new builder.ThumbnailCard(session)
              				.title("OfficeInsights Support")
               			        .subtitle("Please submit your request at //officeinsights")
                 	     	        .text("")
                  		        .images([
              			               builder.CardImage.create(session, "http://officeinsights/Content/images/logo_new.png")
               			         ])
                		        .tap(builder.CardAction.openUrl(session, "https://officemarketinginsights.zendesk.com/hc/en-us"))
           		]);
		session.send(msg);
		session.endDialog();
	}
]);	


dialog.matches('Knowledge', [
	function(session, args, next) {
		var options = {
  		 	uri: 'https://westus.api.cognitive.microsoft.com/qnamaker/v1.0/knowledgebases/2feae40d-56d8-43fe-8cdd-cc20e44cf3af/generateAnswer',
  		 	headers: {
				'Ocp-Apim-Subscription-Key': "be9ab431db0a4e3ca9ec8a12bfc28b8c",
   				'Content-type': 'application/json'
			},
  			 body: {'question': session.message.text},
  			 json: true
		};
		// ====================================
		// Post Requests
		// ====================================
		request.post(options, function (error, response, body) {
			if (!error) {
			   session.send(body.answer);
			} else {
			   session.send(error.message);
			}	
			   session.endDialog();
		}); 
	}
]);


dialog.matches('Profile', [
	function(session, args, next) {
		var s1 = ["smb" ,"ca", "epg" ,"corporate" ,"enterprise", "direct", "reseller", "advisor", "vl","dormant", "initialized", "core", "diversified", "optimized", "e3", "e5", "premium", "essentials"];
		
		words = session.message.text.toLowerCase().replace(",", "").replace(".", "").replace("?", "").split(" ");
		var flg = 0;
		for (var i in words) {	
			if (s1.indexOf(words[i]) != -1){
				flg = 1;
				outstr = "Only Subsidiary and Company views are currently supported. For Segments, Channel or Product info, please visit [//officeinsights](//officeinsights)";
				session.send(outstr);
			}			
		}
		
		if (flg == 0) {
			var country = ["africa new markets","albania","algeria","angola","argentina","armenia","australia","austria","azerbaijan","bahrain","bangladesh","bcbbj","belarus","belgium","bolivia","bosnia and herzegovina","brazil","brunei","bulgaria","cambodia","canada","central asia","chile","china","colombia","costa rica","croatia","cyprus","czech republic","denmark","dominican republic","ecuador","egypt","el salvador","estonia","finland","france","georgia","germany","ghana","greece","guatemala","honduras","hong kong","hungary","iceland","india sc","indian ocean islands","indonesia","iraq","ireland","israel","italy","japan","jordan","kazakhstan","kenya","korea","kuwait","latvia","lebanon","libya","lithuania","luxembourg","macedonia (fyrom)","malaysia","malta","mexico","moldova","montenegro","morocco","myanmar","nepa new markets","netherlands","new zealand","nigeria","norway","oman","pakistan","panama","paraguay","peru","philippines","poland","portugal","puerto rico","qatar","rest of east & southern africa","romania","russia","saudi arabia","serbia","singapore","slovakia","slovenia","south africa","south east asia multi country","spain","sri lanka","sweden","switzerland","taiwan","thailand","trinidad & tobago","tunisia","turkey","turkmenistan","ukraine","united arab emirates","united kingdom","united states","uruguay","venezuela","vietnam","west and central africa"]; 


			var ask = (" " + session.message.text.toLowerCase() + " ");
			ask = ask.replace(",", "").replace(".", "").replace("?", "").replace(" us ", " united states "); 
			console.log(ask);
			var target = "";
			for (var i=0; i < country.length; i++) {	
				if (ask.indexOf(country[i]) >= 0) {
					target = country[i];
				}
			}

			if (target != "") {
				const sql = require("mssql/msnodesqlv8");
				const conn = new sql.ConnectionPool({
 	 			    database: "Commercial",
				    server: "ODNISQL109",
  				    driver: "msnodesqlv8",
 		 		    options: {
  					  trustedConnection: true
 		 			}
				});

				var sque = "select *, round(EOP,0) as EOPx, round(Retention/cast(Volume as float),3) as retentrate from dbo.bizhealth where CountryName = \'" + target + "\'";
				console.log(sque);
				conn.connect().then(function() {
					var request = new sql.Request(conn);
					session.send("Please wait. The query might take a few minutes...");
					session.sendTyping();
					request.query(sque).then(function (recordSet) {
						var arr = recordSet["recordset"];
						var obj = arr[0];
						var outstr = "";
						//outstr += "**" + target.toUpperCase() + "**\n\n";
						outstr += "**Trial**\n\n";
						outstr += "Volume: " + obj["TrialVolumne"] + "\n\n";
						outstr += "Conversion: " + obj["TrialConversions"] + "\n\n";
						outstr += "___________________________\n\n";
		
						outstr += "**Buy**\n\n";
						outstr += "EOP Paid Seats: " + obj["EOPx"] + "\n\n";
						outstr += "Net Paid Seats Add: " + obj["NPSA"] + "\n\n";
						//outstr += "YoY increase: \n\n";
						outstr += "___________________________\n\n";
		
						outstr += "**Use**\n\n";
						outstr += "EXO QE: " + obj["EXO_QE"] + "\n\n";
						outstr += "EXO AE: " + obj["EXO_AE"] + "\n\n";
						outstr += "SfB QE: " + obj["SfB_QE"] + "\n\n";
						outstr += "SfB AE: " + obj["SfB_AE"] + "\n\n";
						outstr += "SPO QE: " + obj["SPO_QE"] + "\n\n";
						outstr += "SPO AE: " + obj["SPO_AE"] + "\n\n";
						outstr += "Yammer QE: " + obj["Yammer_QE"] + "\n\n";
						outstr += "Yammer AE: " + obj["Yammer_AE"] + "\n\n";
						outstr += "___________________________\n\n";

						outstr += "**Renew**\n\n";
						outstr += "Cohort Volume: " + obj["Volume"] + "\n\n";
						outstr += "Renew Rate: " + obj["retentrate"] + "\n\n";
						session.sendTyping();
						session.send(outstr); 	
						
					}).catch(function (err) {
						console.log(err);
						conn.close();
					});
				});
			} else {
				session.beginDialog('companyProfile');
				//session.endDialog();
			}
						
		}
	}
]);



bot.dialog('companyProfile', [
	function (session, results, next) {
		var lexstr = session.message.text.toLowerCase().replace("show", "").replace("usage", "").replace("profile","");
		var words = new pos.Lexer().lex(lexstr);
		var tagger = new pos.Tagger();
		var taggedWords = tagger.tag(words);
		console.log(taggedWords);

		var tenantname = "";
		for (i in taggedWords) {
			var tagword = taggedWords[i];
			if (tagword[1] == 'NNP' || tagword[1] == 'NN') {
				tenantname += tagword[0] + " ";
			}
		}
		console.log(tenantname);


		session.send('Please wait. The results may take a few minutes...');
		const sql = require("mssql/msnodesqlv8");
		const conn = new sql.ConnectionPool({
 	 	  database: "Commercial",
		  server: "ODNISQL109",
  		  driver: "msnodesqlv8",
 		  options: {
  			  trustedConnection: true
 		 	}
		});

		conn.connect().then(function() {
			var request = new sql.Request(conn);
			var str = "SELECT * FROM bbProfile(" + "\'" + tenantname.trim() + "\'" + ")";
			session.sendTyping(); 
			console.log(str);
			request.query(str).then(function (recordSet) {
				var arr = recordSet["recordset"];
				console.log(arr.length);
				if (arr.length > 20) {
					session.send("Too many matches. Please try with another name or search by TenantId.");
					session.endDialog();
					//session.beginDialog('otherCo');
				} else if (arr.length == 0) {
					session.send("No match found. Please try with another name or search by TenantId.");
					session.endDialog();
					//session.beginDialog('otherCo');
				} else {
					session.dialogData.queryresults = arr;
					session.sendTyping(); 
					if (arr.length > 1) {
						var elem = [];
						arr.forEach( function(item) {
							elem.push(item["OrganizationName"] + "  (Total_AvailableUnits: " + item["Total_AvailableUnits"] + ")");
						});
						builder.Prompts.choice(session, "Multiple Matches found. Select the one that you are interested in:\n\n", elem); 	
					} else {
						next();
					}
				}
			}).catch(function (err) {
				console.log(err);
				conn.close();
			});
		});
	},
	function (session, results) {
		var arr = session.dialogData.queryresults;
		var obj = [];
		
		if (arr.length > 1) {	
			obj = arr[results.response.index];
		} else {
			obj = arr[0];
		}

		var outstr = "";
		outstr += "**" + obj["OrganizationName"] + " (" + obj["CountryCode"] +") ** " + "\n\n";
		outstr += "**" + obj["TenantId"] + "**  \n\n";
		outstr += "________________________________________\n\n";
		outstr += "*Category*: `" + obj["TenantCategory_Snapshot"] + "` *Seat Size* : `" + obj["Total_AvailableUnits"] + "`\n\n";
		outstr += "*Tenant_Tenure*: `" + obj["Tenant_Tenure"] + "` *Channel* : `" + obj["DetailSalesModel"] + "`\n\n";
		outstr += "*AOM Segment*: `" + obj["Usage_Segment"] + "`\n\n";
		//outstr += "*Seat Size* : " + obj["Total_AvailableUnits"] + "\n\n";
		outstr += "*[EXO Users]* " + " Enabled: `" + obj["EXOEnabledUsers"] + "` ... Active: `" + obj["EXOActUserslst28"] + "` ... Mobile: `" + obj["MobileEXOActvUsersL28"] + "`\n\n";
		outstr += "*[SfB Users]* " + " Enabled: `" + obj["LYOEnabledUsers"] + "` ... Active: `" + obj["LYOActUserslst28"] + "` ... Mobile: `" + obj["MobileSfBActvL28"] + "`\n\n"; 	
		outstr += "*[SPO Users]* " + " Enabled: `" + obj["SPOEnabledUsers"] + "` ... Active: `" + obj["SPOActUserslst28"] + "` ... Mobile: `" + obj["MobileSPOActvUsersL28"] + "`\n\n";
		outstr += "*[ODB Users]* " + " Enabled: `" + obj["OD4BEnabledUsers"] + "` ...Active: `" + obj["OD4BActUserslst28"] + "`\n\n";
		outstr += "*[Yammer Users]*   " + "Enabled: `" + obj["YammerEnabledUsers"] + "` ... Active: `" + obj["YammerActUserslst28"] + "`\n\n";
		outstr += "*[ProPlus Users]*   " + "Deployed: `" + obj["PPDDeployedUsers"] + "`\n\n";
		//outstr += "*Channel* : `" + obj["DetailSalesModel"] + "`\n\n";
		//outstr += "*AOM Segment* : `" + obj["Usage_Segment"] + "`\n\n";
		outstr += "*Recommendation* : " + obj["NextBestActionName"] + "\n\n";  

		session.send(outstr);
		session.endDialog();
	}
]);



dialog.matches('TopTenants', [
	function (session, results, next) {
		var segm = []; 
		var channel = [];
		var useg = [];
		var prod = [];

		var s1 = ["smb" ,"ca", "epg" ,"corporate" ,"enterprise"];
		var s2 = ["direct", "reseller", "advisor", "vl"];
		var s3 = ["dormant", "initialized", "core", "diversified", "optimized"];
		var s4 = ["e3", "e5", "premium", "essentials"];

		var words = session.message.text.toLowerCase().replace(",", "").replace(".", "").replace("?", "").replace("enterprise", "epg").replace("corporate", "ca").split(" "); 
		for (var i in words) {	
			if (s1.indexOf(words[i]) != -1){
				segm.push("\'" + words[i].toUpperCase() + "\'");
				segm.push(",");
			}
			if (s2.indexOf(words[i]) != -1){
				channel.push("\'" + words[i].toUpperCase() + "\'");
				channel.push(",");
			}
			if (s3.indexOf(words[i]) != -1){
				useg.push("\'" + words[i].toUpperCase() + "\'");
				useg.push(",");
			}
			if (s4.indexOf(words[i]) != -1){
				prod.push("OfferProductName like \'%" + words[i].toUpperCase() + "%\'");
				prod.push(" OR ");
			}	
		}
	
		var str1 = "";
		var str2 = "";
		var str3 = "";
		var str4 = ""; 

		if (segm.length != 0){
			str1 = " AND FinalSegment IN (";
			segm.pop();
			segm.push(")");
			for (var j in segm) {
				console.log(j);
	    		        str1 += segm[j];
			}
			console.log(str1);	
		} 
		if (channel.length != 0){
			str2 = " AND DetailSalesModel IN (";
			channel.pop();
			channel.push(")");
			for (var j in channel) {
				str2 += channel[j];
			}
			console.log(str2);
		} 
		if (useg.length != 0){
			str3 = " AND Usage_Segment IN (";
			useg.pop();
			useg.push(")");
			for (var j in useg) {
				str3 += useg[j];
			}
			console.log(str3);
		}
		if (prod.length != 0){
			str4 = " AND (";
			prod.pop();
			prod.push(")");
			for (var j in prod) {
				str4 += prod[j];
			}
			console.log(str4);
		}
		 
		var str5 = str1 + str2 + str3 + str4; 
		if (str5 == "") {
			session.endDialog("Don't understand your filters. Oops...");
		} else {
			session.dialogData.squery = str5;
			next();
		}
	},
	function (session) {
		const sql = require("mssql/msnodesqlv8");
		const conn = new sql.ConnectionPool({
 	 	  database: "Commercial",
		  server: "ODNISQL109",
  		  driver: "msnodesqlv8",
 		  options: {
  			  trustedConnection: true
 		 	}
		});

		var sque = "select TOP 10 [OrganizationName],[TenantId],[CountryCode], \'Paid\' AS [TenantCategory_Snapshot], [Total_AvailableUnits], [TotalActUserslst28] from [dbo].[PaidTenantSummary] where 1=1" + session.dialogData.squery + " ORDER BY Total_AvailableUnits DESC";
		console.log(sque);
		conn.connect().then(function() {
			var request = new sql.Request(conn);
			session.send("Please wait. The query might take a few minutes...");
			session.sendTyping();
			request.query(sque).then(function (recordSet) {
				var arr = recordSet["recordset"];
				var outstr = "";
				for (var i = 0; i < arr.length; i++){
					var obj = arr[i];
					
					outstr += "**" + (i+1) + " " + obj["OrganizationName"] + "** \n\n";
					var j = 0;
					for (var key in obj) {	
					   if (j > 0) {
						outstr += " *" + key + "* : " + obj[key] + "\n\n";						
					   }
					   j++;
					}
					if (i < arr.length-1) {
						outstr += "___________________________________\n\n";
					}
				}
				session.sendTyping();
				session.endDialog(outstr); 
			}).catch(function (err) {
				console.log(err);
				conn.close();
			});
		});
	}
]);


