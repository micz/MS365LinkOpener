async function add365LinkButton() {
	const links = extractLinks(document.body.innerHTML);
	//console.log(links);
	//console.log(document.body.innerHTML);
	
	// Create the banner element itself.
    const banner = document.createElement("div");
    banner.className = "miczMS365Opener_btn";

	var links_done = [];

	for(let i=0;i<links.length;i++){
		
		if (links_done.indexOf(links[i].href) !== -1) {
		  	continue;
		}
		links_done.push(links[i].href);
		
		var app_protocol = getFileType(links[i]);
		var app_selector = null;
		//console.log('>>>>>>>>>>>>>> [add365LinkButton] PRE force_msedge');
		let prefs = await browser.storage.sync.get({force_msedge: false, always_link: false});
		//console.log('>>>>>>>>>>>>>> [add365LinkButton] force_msedge: '+prefs.force_msedge );
		if(app_protocol!==''){		// App found
			app_selector = getAppBtn(links[i].href, prefs, {ppt: app_protocol==='p', wrd: app_protocol==='w', xls: app_protocol==='x', lnk: false} );
			banner.appendChild(app_selector);
			
		}else{	//App not found
			app_selector = getAppBtn(links[i].href, prefs);
			banner.appendChild(app_selector);
		}
		let linkText = document.createElement("span");
		linkText.innerText = links[i].text;
		app_selector.appendChild(linkText);
	}

    // Insert it as the very first element in the message.
    document.body.insertBefore(banner, document.body.firstChild);
	//console.log(document.body.innerHTML);
};

function extractLinks(html) {
	  const links = [];
	  const parser = new DOMParser();
	  const doc = parser.parseFromString(html, 'text/html');
	  const anchorElements = doc.querySelectorAll('a');

	  anchorElements.forEach((element) => {
		const href = element.getAttribute('href');
		let text = element.textContent.trim();

		if (text.includes(href)) {
			text = "";
		}
		
		if (href && (!href.includes('/_layouts/') || href.includes('file=')) && !href.includes('/SitePages/') && (href.includes('/sites/') || href.includes('/personal/') || href.match(/:[A-Za-z]:/)) && href.match(/^https?:\/\/[a-zA-Z0-9.-]+\.sharepoint\.com/i)) {
		  links.push({ href, text });
		}
	  });

	  return filterLinks(links);
}

function filterLinks(links) {
	const uniqueHrefs = {};

	links = sortLinksByText(links);
  
	// Filter the array
	const filteredLinks = links.filter(link => {
	  // Se l'href è già presente e il testo è vuoto, ignorare l'elemento
	  if (uniqueHrefs[link.href] && link.text === '') {
		return false;
	  }
  
	  // Segnare l'href come già visto
	  uniqueHrefs[link.href] = true;
  
	  // Includere l'elemento nell'array filtrato
	  return true;
	});
  
	return filteredLinks;
  }

function sortLinksByText(links) {
	// Ordina l'array mettendo dopo gli oggetti con text vuoto
	links.sort((a, b) => {
	  if (a.text === '' && b.text !== '') {
		return 1; // metti 'a' prima di 'b'
	  } else if (a.text !== '' && b.text === '') {
		return -1; // metti 'b' prima di 'a'
	  } else {
		return 0; // lascia invariato l'ordine
	  }
	});
	return links;
  }
  

function getFileType(test){	//ritorna '' se non si è capito quale sia - w se word - x se excel - p se powerpoint
  let regex_xls = /\.(xlsx|xls|ods)/i;
  let is_excel = regex_xls.test(test.href) || regex_xls.test(test.text);
  let regex_doc = /\.(docx|doc|odt)/i;
  let is_word = regex_doc.test(test.href) || regex_doc.test(test.text);
  let regex_ppt = /\.(pptx|ppt|odp)/i;
  let is_powerpoint = regex_ppt.test(test.href) || regex_ppt.test(test.text);//console.log('is_powerpoint: '+is_powerpoint);
  //if (is_excel===is_word || is_excel===is_powerpoint || is_word===is_powerpoint) return '';
  let checkApp = (is_word ? 1 : 0) + (is_excel ? 1 : 0) + (is_powerpoint ? 1 : 0);
  if((checkApp == 0) || (checkApp > 1)) return '';
  if (is_excel) return 'x';
  if (is_word) return 'w';
  if (is_powerpoint) return 'p';
}

function getAppBtn(link, prefs = false, par = {ppt: true, wrd: true, xls: true, lnk: true}) {
  // Create the selector
  var app_selector = document.createElement("div");
  app_selector.className = "miczMS365Opener_wrapper";

  if(!prefs){
	prefs = {force_msedge: false, always_link: false};
  }

  if(prefs.always_link){
	par.lnk = true;
  }

  //console.log('>>>>>>>>>>>>>> [getAppBtn] force_msedge: '+force_msedge);

  // Remove che query string "d" parameter, or the app will open an empty document
  // This parameter is used in the links in a comment notification email
  link = removeQueryParam(link,'d');

  // Array of options with their values and texts
  var options = {
    'lnk': { value: (prefs.force_msedge?"microsoft-edge:":"") + link, text: browser.i18n.getMessage("openLink"), image: browser.runtime.getURL('../images/link-32px.png') },
	'ppt': { value: "ms-powerpoint:ofe|u|" + link, text: browser.i18n.getMessage("openPowerPoint"), image: browser.runtime.getURL('../images/powerpoint-32px.png') },
	'xls': { value: "ms-excel:ofe|u|" + link, text: browser.i18n.getMessage("openExcel"), image: browser.runtime.getURL('../images/excel-32px.png') },
    'wrd': { value: "ms-word:ofe|u|" + link, text: browser.i18n.getMessage("openWord"), image: browser.runtime.getURL('../images/word-32px.png') },
	};

  // Add options to the selector
  Object.keys(options).forEach(key => {
	  if(par[key]){
		let option = options[key];
		let btnElement = document.createElement("button");
		//btnElement.innerText = option.text;
		let img = new Image();
		img.src = option.image;
		img.alt = option.text;
		img.title = option.text;
		img.width = 16;
		img.height = 16;
		btnElement.appendChild(img);
		
		let aElement = document.createElement("a");
		aElement.href = option.value;

		aElement.appendChild(btnElement);
		app_selector.appendChild(aElement);
	  }
  });
  
  return app_selector; 
}

function removeQueryParam(url, paramToRemove) {
    const urlObject = new URL(url);
    urlObject.searchParams.delete(paramToRemove);
    return urlObject.toString();
}

add365LinkButton();
