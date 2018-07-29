$(document).ready(function() {
	AddressBook.init(); //initilization of the state of DOM Elements
	AddressBook.bindEvents(); //event handler associated with the Address Book
});

let AddressBook = {
	AddressBookList: 'AddressBook',
	CompanyList: 'Company',
	MaritalStatusField: 'MaritalStatus',
	AddressBookListFields:
		'?$select=Id,FullName,EMail,CellPhone,HomePhone,WebPage,WorkAddress,DateOfBirth,MaritalStatus' +
		'&$select=LookupCompany/Title,LookupCompany/Id,Relative/Title,Relative/EMail,Relative/Id,Browser' +
		'&$expand=LookupCompany/Title,LookupCompany/Id,Relative/Title,Relative/EMail,Relative/Id',
	AddressBookListItemType: 'SP.Data.AddressBookListItem',
	maritalStatusDiv: '#maritalStatus',
	relativeDivId: 'relative', //this is a people picker div
	BrowserTermSetId: '4e1ea6a1-5ab8-4325-8aac-153a5e8badf7', // for TermSet: 'Internet Browser'
	lookupCompanyDiv: '#company',
	RelativesGroupId: 26, //Relatives GroupId: 26
	WssId: -1, //this is for metadata column

	//validation patterns
	usernamePattern: /^([A-Za-z\s]+)$/,
	emailPattern: /^([a-zA-Z0-9\_\-\.]{2,})@(([a-zA-Z0-9\_\-\.]{2,}))\.([a-zA-Z]{2,3})$/,
	urlPattern: /^(https:|http:|ftp:)(\/{2})((([w]{3})\.([a-z]{3,})\.([a-z]{2,3}))|(([a-z]{3,})\.([a-z]{2,3})))$/i,
	mobilePattern: /^\d{10}$/,
	landlinePattern: /^\d{10}$/,
	dateOfBirthPattern: /^((19|20)\d{2})[-]((0?[1-9])|(1[0-2]))[-]((0?[0-9])|([1-2][0-9])|(3[0-1]))$/, // yyyy-mm-dd

	//ajax call methods
	ajaxCallMethod: {
		_get: 'GET',
		post: 'POST'
	},
	ajaxCallHeaders: {
		_getHeader: { accept: 'application/json;odata=verbose' },
		_addHeader: {
			accept: 'application/json;odata=verbose',
			'X-RequestDigest': $('#__REQUESTDIGEST').val(),
			'content-type': 'application/json;odata=verbose'
		},
		_updateHeader: {
			accept: 'application/json;odata=verbose',
			'X-RequestDigest': $('#__REQUESTDIGEST').val(), //"X-RequestDigest": "/_api/contextinfo",  //throw 403() forbidden error
			'content-type': 'application/json;odata=verbose',
			'IF-MATCH': '*',
			'X-HTTP-Method': 'MERGE'
		},
		_deleteHeader: {
			accept: 'application/json;odata=verbose',
			'content-type': 'application/json;odata=verbose',
			'X-RequestDigest': $('#__REQUESTDIGEST').val(),
			'IF-MATCH': '*',
			'X-HTTP-Method': 'DELETE'
		}
	},

	//alertMessages for success methods
	alertMessages: {
		added: 'Contact created Successfully',
		updated: 'Contact updated Successfully',
		deleted: 'Contact deleted Successfully'
	},
	contactDetails: [], //the static array which is act as database
	browsers: [], //it will contain all the terms from the termset: Internet Browser
	Contact: function(obj) {
		this.Id = obj.Id;
		this.FullName = obj.FullName;
		this.DateOfBirth = obj.DateOfBirth; //this is date and time column in the list
		this.MaritalStatus = obj.MaritalStatus; //this is a choice column in the list
		this.LookupCompany = { Id: obj.LookupCompany.Id, Title: obj.LookupCompany.Title }; //this is a lookup column in the list
		this.Browser = { Label: obj.Browser.Label, TermGuid: obj.Browser.TermGuid, WssId: AddressBook.WssId }; //this is a managed metadata column in the list
		this.Relative = { EMail: obj.Relative.EMail, Title: obj.Relative.Title }; //this is person or group column in the list
		this.EMail = obj.EMail;
		(this.CellPhone = obj.CellPhone), //mobile number in form field
			(this.HomePhone = obj.HomePhone), //landline number in form field
			(this.WebPage = { Url: obj.WebPage.Url }), //this is hyperlink column in the list
			(this.WorkAddress = obj.WorkAddress); //this is 'address' in the address form field
	},

	init: function() {
		$('.form-container, .contact-info, div#errorMessage').hide();
		$('input').attr('autocomplete', 'off');

		//initialize the people picker div element
		AddressBook.initializePeoplePicker(AddressBook.relativeDivId, AddressBook.RelativesGroupId);

		//get and bind all choices
		$.when(AddressBook.getAllChoices(AddressBook.AddressBookList, AddressBook.MaritalStatusField)).done(function(
			fieldChoices
		) {
			AddressBook.bindAllChoices(fieldChoices, AddressBook.maritalStatusDiv);
		});

		//get and bind all lookupValues
		$.when(AddressBook.getLookupValues(AddressBook.CompanyList)).done(function(lookupValues) {
			AddressBook.bindLookupValues(lookupValues, AddressBook.lookupCompanyDiv);
		});

		//get and bind all terms: these are the metadata terms from the termset 'Internet Browser'
		$.when(AddressBook.getAllTerms(AddressBook.BrowserTermSetId)).done(function(allTerms) {
			AddressBook.bindAllTerms(allTerms);
		});

		//get all the contacts from the sharepoint list: AddressBook and bind all Contacts
		$.when(AddressBook.getAllContacts(AddressBook.AddressBookList, AddressBook.AddressBookListFields)).done(
			function(contacts) {
				if (contacts.length > 0) {
					$.each(contacts, function(index, obj) {
						let Browser = {};

						//this is required in case if single value is allowed
						//it is not required in case if multiple values are allowed

						$.each(AddressBook.browsers, function(browserIndex, browserObject) {
							//first: check if the browser data is in the list or not
							if (obj.Browser != null && obj.Browser != '' && obj.Browser != undefined) {
								if (obj.Browser.TermGuid == browserObject.id) {
									Browser = {
										Label: browserObject.Title,
										TermGuid: obj.Browser.TermGuid,
										WssId: AddressBook.WssId
									};
									return false; //break out of each inner loop
								}
							}
						});

						const contact = new AddressBook.Contact({
							Id: obj.Id,
							FullName: obj.FullName,
							DateOfBirth:
								obj.DateOfBirth != null
									? obj.DateOfBirth.substring(0, obj.DateOfBirth.indexOf('T'))
									: null,
							MaritalStatus: obj.MaritalStatus,
							LookupCompany:
								obj.LookupCompany != null
									? { Id: obj.LookupCompany.Id, Title: obj.LookupCompany.Title }
									: null,
							Browser: Browser,
							Relative:
								obj.Relative != null ? { Title: obj.Relative.Title, EMail: obj.Relative.EMail } : null,
							EMail: obj.EMail,
							CellPhone: obj.CellPhone,
							HomePhone: obj.HomePhone,
							WebPage: obj.WebPage != null ? { Url: obj.WebPage.Url } : null,
							WorkAddress: obj.WorkAddress
						});

						//adding to the contactDetails array
						AddressBook.contactDetails.push(contact);
					});

					//bind all the contacts to the webpage
					AddressBook.bindContacts(AddressBook.contactDetails);

					//hide 'no contact' label
					$('#noContacts')
						.parent()
						.hide();
				} else {
					$('#noContacts')
						.parent()
						.show();
				}
			}
		);
	},

	initializePeoplePicker: function(peoplePickerDivId, sharepointGroupId) {
		// Create a schema to store picker properties, and set the properties.
		let schema = {};
		schema['PrincipalAccountType'] = 'User';
		schema['SearchPrincipalSource'] = 15; //This value specifies where you would want to search for the valid values
		schema['ResolvePrincipalSource'] = 15; //This value specifies where you would want to resolve for the valid values
		schema['AllowMultipleValues'] = false; // you can define single/multiple here
		schema['MaximumEntitySuggestions'] = 50;
		schema['SharePointGroupID'] = sharepointGroupId; //here, the sharepoint group id: 26

		SPClientPeoplePicker_InitStandaloneControlWrapper(peoplePickerDivId, null, schema);
	},

	getAllChoices: function(listName, choiceField) {
		const deferred = $.Deferred();
		$.ajax({
			url:
				_spPageContextInfo.webAbsoluteUrl +
				'/_api/web/lists/getbytitle' +
				"('" +
				listName +
				"')" +
				'/fields/getbytitle' +
				"('" +
				choiceField +
				"')" +
				'/Choices',
			method: AddressBook.ajaxCallMethod._get,
			headers: AddressBook.ajaxCallHeaders._getHeader,
			success: function(data) {
				return deferred.resolve(data.d.Choices.results);
			},
			error: AddressBook.error
		});
		return deferred.promise();
	},

	//bind field choices to  marital status form field
	bindAllChoices: function(fieldChoices, choiceFieldDiv) {
		$.each(fieldChoices, function(index, value) {
			let choice = '';
			choice += "<option value ='" + value + "'>" + value + '</option>';

			$(choiceFieldDiv).append(choice);
		});
	},

	//get all lookup companies of the company form field
	getLookupValues: function(lookupListName) {
		const deferred = $.Deferred();
		$.ajax({
			url:
				_spPageContextInfo.webAbsoluteUrl +
				'/_api/web/lists/getbytitle' +
				"('" +
				lookupListName +
				"')" +
				'/items?$Select=Id,Title',
			method: AddressBook.ajaxCallMethod._get,
			headers: AddressBook.ajaxCallHeaders._getHeader,
			success: function(data) {
				return deferred.resolve(data.d.results);
			},
			error: AddressBook.error
		});
		return deferred.promise();
	},

	//bind all lookup companies to the company form field
	bindLookupValues: function(lookupValues, lookupDiv) {
		$.each(lookupValues, function(index, obj) {
			let value = '';
			value += "<option value ='" + obj.Title + "' id = '" + obj.Id + "'>" + obj.Title + '</option>';
			$(lookupDiv).append(value);
		});
	},

	//these terms  will contain "browsers" from the termset: "Internet Browser"
	getAllTerms: function(termSetId) {
		const deferred = $.Deferred();
		const context = SP.ClientContext.get_current();
		const taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
		const termStore = taxonomySession.getDefaultSiteCollectionTermStore();
		const termSet = termStore.getTermSet(termSetId); // here, TermSet id is of TermSet: 'Internet Browser'
		const terms = termSet.getAllTerms();

		context.load(terms);
		context.executeQueryAsync(
			Function.createDelegate(this, function() {
				return deferred.resolve(terms);
			}),
			Function.createDelegate(this, function() {
				AddressBook.error(errorMessage);
			})
		);
		return deferred.promise();
	},

	//it will bind the "browsers" terms
	bindAllTerms: function(terms) {
		let termItems = terms.getEnumerator();
		while (termItems.moveNext()) {
			let term = termItems.get_current();
			let browser = '';
			browser +=
				"<option value ='" +
				term.get_name().toString() +
				"' id = '" +
				term.get_id().toString() +
				"'>" +
				term.get_name().toString() +
				'</option>';

			AddressBook.browsers.push({ id: term.get_id(), Title: term.get_name() }); //pushing the terms: internet browser
			$('select#browser').append(browser); //bind the term to the select element with id: browser
		}
	},

	//get all the contacts from the sharepoint list: here, 'AddressBook'
	getAllContacts: function(listName, fields) {
		const deferred = $.Deferred();
		$.ajax({
			url:
				_spPageContextInfo.webAbsoluteUrl +
				"/_api/web/lists/getbytitle('" +
				listName +
				"')" +
				'/items' +
				fields,
			method: AddressBook.ajaxCallMethod._get,
			headers: AddressBook.ajaxCallHeaders._getHeader,
			success: function(data) {
				return deferred.resolve(data.d.results);
			},
			error: AddressBook.eror
		});
		return deferred.promise();
	},

	// Get the user id
	getUserId: function(loginName) {
		const deferred = $.Deferred();
		$.ajax({
			url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/ensureUser('" + encodeURIComponent(loginName) + "')",
			method: AddressBook.ajaxCallMethod.post, // the 'ensureUser' rest api call only works for POST method
			headers: AddressBook.ajaxCallHeaders._addHeader,
			success: function(data) {
				return deferred.resolve(data.d.Id);
			},
			error: function(error) {
				console.log(error); //this is the actual error that can be refer in console
				return deferred.resolve(null); // this implies the user specified does not exist, so we resolve it as null
			}
		});
		return deferred.promise();
	},

	bindContacts: function(Contacts) {
		$.each(Contacts, function(index, obj) {
			let contactData = '';
			contactData +=
				"<div class = 'contact-card'" +
				"contactid = ' " +
				obj.Id +
				"'>" + // "class = 'active-contact'"  will be added when you select the contact
				"<div class = 'contact-nav'>" +
				'<ul>' +
				"<li class = 'name'>" +
				obj.FullName +
				'</li>' +
				"<li class = 'email'>" +
				obj.EMail +
				'</li>' +
				"<li class = 'mobile-number'>" +
				'+ 91 ' +
				obj.CellPhone +
				'</li>' +
				'</ul>' +
				'</div>' +
				'</div>';
			$('.contact-list').append(contactData);
		});

		/**this is to sort a contact list in the alphabetic order
		 *you can make it a separate function
		 * */
		const $div = $('div.contact-card');
		const sortedDiv = $div.sort(function(a, b) {
			return (
				$(a)
					.find('.name')
					.text() >
				$(b)
					.find('.name')
					.text()
			);
		});

		$('.contact-list').html(sortedDiv);
	},

	bindEvents: function() {
		//show the form container when +ADD is clicked in the navigation bar
		$('.navigation-bar').on('click', '#addContact', function() {
			$('.contact-info, #updateBtn').hide();
			$('.form-container, #addBtn').show();
			$('div').removeClass('active-contact');
			AddressBook.clearForm();
		});

		//close the form when the close mark is clicked
		$('.form-container').on('click', '#closeForm', function() {
			//check whether there exists a contact or not
			if ($('.contact-list').children().length == 1) {
				$('#noContacts')
					.parent()
					.show();
			} else if (
				parseInt($('#closeForm').attr('contactid')) ==
				parseInt(
					$('.contact-list')
						.children('.active-contact')
						.attr('contactid')
				)
			) {
				$('.contact-info').show();
			}
			$('.form-container').hide();
			AddressBook.clearForm();
		});

		//add the contact to the address Book
		$('.form-container #contactForm').on('click', '#addBtn', function() {
			AddressBook.addOrUpdateContact($(this).attr('id'));
		});

		//contact form validation: Real Time Validation
		$('input').on('keyup', function() {
			AddressBook.validate($(this));
		});

		//show the person details in detail when the contact list is clicked
		$('.contact-list').on('click', '.contact-card', function() {
			const contactId = parseInt($(this).attr('contactid'));
			const selectedContact = $(this);

			$('.form-container').hide();

			//if the contact list is not empty
			if ($('.contact-list').children().length > 1) {
				$('div').removeClass('active-contact'); //remove the previous highlight of the contact list
				$(this).addClass('active-contact'); // highlight the selected contact details list

				//show the information of corresponding person which is clicked
				AddressBook.showContactDetails(selectedContact);
				$('.contact-info').show();
			}
		});

		//edit the contact details of corresponding person
		$('.edit-delete').on('click', '.editIcon', function() {
			const contactId = parseInt($('.editIcon').attr('contactid'));
			$('.form-container, #updateBtn').show();
			$('#addBtn, .contact-info, div#errorMessage').hide();

			AddressBook.getContactDetails(contactId); //get the contact Details from the contact Details []
		});

		//update the contact of the address Book
		$('.form-container #contactForm').on('click', '#updateBtn', function() {
			$('div#errorMessage').hide();
			AddressBook.addOrUpdateContact($(this).attr('id'));
		});

		//delete the person contact details from address book
		$('.edit-delete').on('click', '.deleteIcon', function() {
			if (confirm('Are you sure you want to delete?')) {
				const contactId = parseInt($('.deleteIcon').attr('contactid'));
				const selectedContact = $('div.active-contact');

				$('.form-container').hide();
				$('.contact-details')
					.removeAttr('contactid')
					.children()
					.empty();

				AddressBook.deleteContact(contactId, selectedContact);
				selectedContact.remove(); //remove the selected contact from the contact list in the webpage
				$('div.editIcon, div.deleteIcon ').removeAttr('contactid'); //remove the corresponding person id as well
				$('div.editIcon, div.deleteIcon').hide();

				//check whether there exists a contact or not
				if ($('.contact-list').children().length == 1) {
					// here the length of '1' exist because, 'no contact' is the hidden div
					$('#noContacts')
						.parent()
						.show();
					$('#updateBtn').removeAttr('contactid'); //remove the previous contactId associated with the updateBtn
				}
			}
		});
	},

	//this will perform, add, update and delete based on the parameters passed
	addOrUpdateOrDeleteAjaxCall: function(url, method, headers, alertMessage, data = null) {
		const deferred = $.Deferred();
		$.ajax({
			url: url,
			method: method,
			data: data,
			headers: headers,
			success: function(data) {
				alert(alertMessage);
				if (data != null && data != undefined && data != '') {
					return deferred.resolve(data.d.Id); //this will fetch the id of newly created contact
				}
			},
			error: AddressBook.error
		});
		return deferred.promise();
	},

	/**TO_DO
	 * define your own error messages
	 * here the error message is generic*
	 * */
	error: function(errorMessage) {
		alert('error occured,see the console');
		console.log(errorMessage);
	},

	addOrUpdateContact: function(addOrUpdateDivId) {
		let isValid = false;
		isValid = AddressBook.validateOnAddUpdate();

		if (isValid) {
			let Relative = {},
				loginName = null,
				LookupCompany = {},
				Browser = {};

			const picker = SPClientPeoplePicker.SPClientPeoplePickerDict[AddressBook.relativeDivId + '_TopSpan']; // [relative is the name of the "people picker div"]
			if (picker.UnresolvedUserCount == 0 && picker.TotalUserCount == 1 /*&& !picker.IsEmpty()*/) {
				Relative.EMail = picker.GetAllUserInfo()[0].EntityData.Email;
				Relative.Title = picker.GetResolvedUsersAsText();
				loginName = picker.GetAllUserInfo()[0].Key;
			}

			if (
				$('select#company')
					.children(':selected')
					.val() != ''
			) {
				LookupCompany.Id = $('select#company')
					.children(':selected')
					.attr('id');
				LookupCompany.Title = $('select#company')
					.children(':selected')
					.val();
			}

			if (
				$('select#browser')
					.children(':selected')
					.val() != null &&
				$('select#browser')
					.children(':selected')
					.val() != ''
			) {
				Browser = {
					Label: $('select#browser')
						.children(':selected')
						.val(),
					TermGuid: $('select#browser')
						.children(':selected')
						.attr('id'),
					WssId: AddressBook.WssId
				};
			}

			if (addOrUpdateDivId === 'addBtn') {
				const newContact = new AddressBook.Contact({
					FullName: $('#username').val(),
					DateOfBirth: $('#dateOfBirth').val(),
					MaritalStatus: $('select#maritalStatus')
						.children(':selected')
						.val(),
					LookupCompany: LookupCompany,
					Browser: Browser,
					Relative: Relative,
					EMail: $('#email').val(),
					CellPhone: $('#mobileNo').val(),
					HomePhone: $('#landlineNo').val(),
					WebPage: { Url: $('#url').val() },
					WorkAddress: $('#addressDetails').val()
				});

				$.when(AddressBook.getUserId(loginName))
					.then(function(userId) {
						if ($.isEmptyObject(Browser)) {
							Browser = null; //this is the explicit value that is passed to the managed metadata column( which implies the read mode i.e. no new value is assigned or added to the field)
						}
						const data = Object.assign(
							{},
							{ __metadata: { type: AddressBook.AddressBookListItemType } },
							{ RelativeId: userId },
							newContact
						);
						data.Browser = Browser;

						const url =
							_spPageContextInfo.webAbsoluteUrl +
							'/_api/web/lists/getbytitle' +
							"('" +
							AddressBook.AddressBookList +
							"')" +
							'/items';
						//adding new contact to sharepoint list
						return AddressBook.addOrUpdateOrDeleteAjaxCall(
							url,
							AddressBook.ajaxCallMethod.post,
							AddressBook.ajaxCallHeaders._addHeader,
							AddressBook.alertMessages.added,
							JSON.stringify(data)
						);
					})
					.done(function(contactId) {
						newContact.Id = contactId; //retrieving the newly created contact id
						AddressBook.contactDetails.push(newContact); //adding new contact detail to the contact Details Array

						let contact = [];
						contact.push(newContact);
						AddressBook.bindContacts(contact); //function to bind the person data to the webpage

						$('#closeForm').attr('contactid', newContact.Id); //associate the contact id with close form
					});
			} else if (addOrUpdateDivId === 'updateBtn') {
				//alert("update");
				const contactid = parseInt($('div.contact-card.active-contact').attr('contactid'));
				const selectedContact = $('div.contact-card.active-contact');
				const contact = AddressBook.contactDetails.find(obj => obj.Id === contactid); //find the contact from the contact Details Array

				//update the correspoonding contact details in the contactDetails Array
				contact.FullName = $('#username').val();
				contact.DateOfBirth = $('#dateOfBirth').val();
				contact.MaritalStatus = $('select#maritalStatus')
					.children(':selected')
					.val();
				contact.LookupCompany = LookupCompany;
				contact.Browser = Browser;
				contact.Relative = Relative;
				contact.EMail = $('#email').val();
				contact.CellPhone = $('#mobileNo').val();
				contact.HomePhone = $('#landlineNo').val();
				contact.WebPage.Url = $('#url').val();
				contact.WorkAddress = $('#addressDetails').val();

				//first, get the user id of the people picker,
				//then update the conact in the Sharepoint list
				$.when(AddressBook.getUserId(loginName))
					.then(function(userId) {
						if ($.isEmptyObject(Browser)) {
							Browser = null; //this is the explicit value that is passed to the managed metadata column( which implies the read mode i.e. no new value is assigned or added to the field)
						}
						const data = Object.assign(
							{},
							{ __metadata: { type: AddressBook.AddressBookListItemType } },
							{ RelativeId: userId },
							contact
						);
						data.Browser = Browser;

						const url =
							_spPageContextInfo.webAbsoluteUrl +
							'/_api/web/lists/getbytitle' +
							"('" +
							AddressBook.AddressBookList +
							"')" +
							'/items' +
							'(' +
							contact.Id +
							')';

						//then update the list in the sharepoint
						AddressBook.addOrUpdateOrDeleteAjaxCall(
							url,
							AddressBook.ajaxCallMethod.post,
							AddressBook.ajaxCallHeaders._updateHeader,
							AddressBook.alertMessages.updated,
							JSON.stringify(data)
						);
					})
					.done(function() {
						//finally, show the information of corresponding person which is clicked
						selectedContact.find('.name').text(contact.FullName);
						selectedContact.find('.email').text(contact.EMail);
						selectedContact.find('.mobile-number').text('+ 91 ' + contact.CellPhone);

						AddressBook.showContactDetails(selectedContact);
						$('.contact-info').show();
					});
			}

			$('.form-container, #noContacts').hide();
			AddressBook.clearForm();
		} else {
			$('div#errorMessage')
				.text('Please Check Your Input')
				.show();
		}
	},

	deleteContact: function(contactid, selectedContact) {
		const url =
			_spPageContextInfo.webAbsoluteUrl +
			'/_api/web/lists/getbytitle' +
			"('" +
			AddressBook.AddressBookList +
			"')" +
			'/items' +
			'(' +
			contactid +
			')';

		//delete the contact from the sharepoint list and from the contact details [] array
		AddressBook.addOrUpdateOrDeleteAjaxCall(
			url,
			AddressBook.ajaxCallMethod.post,
			AddressBook.ajaxCallHeaders._deleteHeader,
			AddressBook.alertMessages.deleted
		);
		AddressBook.contactDetails = AddressBook.contactDetails.filter(obj => obj.id !== contactid);
	},

	//this will populate the data in the form field to edit
	getContactDetails: function(contactid) {
		//first, find the contact in the contact Details []
		const contact = AddressBook.contactDetails.find(obj => obj.Id === contactid);

		//resolving the user for the people picker field, it will populate the data automatically for the people picker field
		if (contact.Relative.Title != null && contact.Relative.Title != '' && contact.Relative.Title != undefined) {
			const picker = SPClientPeoplePicker.SPClientPeoplePickerDict[AddressBook.relativeDivId + '_TopSpan']; // [relative is the name of the "people picker div"]
			picker.DeleteProcessedUser(); //clear the previous populated relative name
			const userObj = { Key: contact.Relative.EMail };
			picker.AddUnresolvedUser(userObj, true); //add the selected contact "relative"
		}

		//then populate the remaining form field
		$('#username').val(contact.FullName);
		$('#dateOfBirth').val(contact.DateOfBirth);
		$('select#maritalStatus').val(contact.MaritalStatus);
		$('select#company').val(contact.LookupCompany.Title);
		$('select#browser').val(contact.Browser.Label);
		$('#email').val(contact.EMail);
		$('#mobileNo').val(contact.CellPhone);
		$('#landlineNo').val(contact.HomePhone);
		$('#url').val(contact.WebPage.Url);
		$('#addressDetails').val(contact.WorkAddress);
	},

	//this will show all information related to the contact
	showContactDetails: function(selectedContact) {
		const contactid = parseInt(selectedContact.attr('contactid'));
		const contact = AddressBook.contactDetails.find(obj => obj.Id === contactid); //find the contact and show the details (this method will return undefined in case contact is not found)

		//then fill the contact details
		$('#personName').text(AddressBook.isPropertyExist(contact.FullName) ? contact.FullName : '');
		$('#personDateOfBirth').text(
			AddressBook.isPropertyExist(contact.DateOfBirth) ? 'Date of Birth: ' + contact.DateOfBirth : ''
		);
		$('#personMaritalStatus').text(
			AddressBook.isPropertyExist(contact.MaritalStatus) ? 'Marital Status: ' + contact.MaritalStatus : ''
		);
		$('#personCompany').text(
			AddressBook.isPropertyExist(contact.LookupCompany.Title)
				? 'Company Name: ' + contact.LookupCompany.Title
				: ''
		);
		$('#browserName').text(
			AddressBook.isPropertyExist(contact.Browser.Label) ? 'Browser Name: ' + contact.Browser.Label : ''
		);
		$('#personRelative').text(
			AddressBook.isPropertyExist(contact.Relative.Title) ? 'Relative Name: ' + contact.Relative.Title : ''
		);
		$('#emailId').text(AddressBook.isPropertyExist(contact.EMail) ? 'Email: ' + contact.EMail : '');
		$('#mobileNumber').text(AddressBook.isPropertyExist(contact.CellPhone) ? 'Mobile: ' + contact.CellPhone : '');
		$('#landlineNumber').text(
			AddressBook.isPropertyExist(contact.HomePhone) ? 'Landline: ' + contact.HomePhone : ''
		);
		$('#website').text(
			AddressBook.isPropertyExist(contact.WebPage.Url) ? 'Website: ' + contact.WebPage.Url.toLowerCase() : ''
		);
		$('#address').html(
			AddressBook.isPropertyExist(contact.WorkAddress)
				? 'Address: ' + contact.WorkAddress.replace(/\n/g, '<br>')
				: ''
		); //fomatting address for UI

		$('div.contact-details').attr('contactid', contact.Id); //associate the id with the corresponding contact
		$('div.editIcon, div.deleteIcon ,#updateBtn, #closeForm').attr('contactid', contact.Id); //binding edit,delete,update and closeBtn icon to the particular contact in the contact list
		$('div.editIcon, div.deleteIcon').show(); //show the edit and delete icon
	},

	isPropertyExist: function(propertyValue) {
		return propertyValue !== null && propertyValue !== '' && propertyValue !== undefined ? true : false;
	},

	//validation on Real Time i.e when the user enter the value in the input field
	validate: function(inputField) {
		$('#errorMessage').hide();

		//checking for validation against each required input field
		switch (inputField.attr('id')) {
			case 'username':
				AddressBook.validityMessage(inputField, AddressBook.usernamePattern);
				break;
			case 'dateOfBirth':
				AddressBook.validityMessage(inputField, AddressBook.dateOfBirthPattern);
				break;
			case 'email':
				AddressBook.validityMessage(inputField, AddressBook.emailPattern);
				break;
			case 'url':
				AddressBook.validityMessage(inputField, AddressBook.urlPattern);
				break;
			case 'mobileNo':
				AddressBook.validityMessage(inputField, AddressBook.mobilePattern);
				break;
			case 'landlineNo':
				AddressBook.validityMessage(inputField, AddressBook.landlinePattern);
				break;
			case '':
				inputField
					.parent()
					.siblings('.error')
					.text('cannot be Empty');
				inputField
					.parent()
					.siblings('.error')
					.show();
				break;
		}
	},

	validityMessage: function(inputField, inputFieldPattern) {
		if (!inputFieldPattern.test(inputField.val())) {
			inputField
				.parent()
				.siblings('.error')
				.text('not valid')
				.show();
		} else {
			inputField
				.parent()
				.siblings('.error')
				.text('');
		}
	},

	validateOnAddUpdate: function() {
		const validity =
			AddressBook.usernamePattern.test($('#username').val()) &&
			AddressBook.dateOfBirthPattern.test($('#dateOfBirth').val()) &&
			AddressBook.emailPattern.test($('#email').val()) &&
			AddressBook.urlPattern.test($('#url').val()) &&
			AddressBook.mobilePattern.test($('#mobileNo').val()) &&
			AddressBook.landlinePattern.test($('#landlineNo').val());
		return validity;
	},

	clearForm: function() {
		$('#contactForm')
			.find('input,textarea,select')
			.val('');
		$('.error, #errorMessage').hide();
		const picker = SPClientPeoplePicker.SPClientPeoplePickerDict[AddressBook.relativeDivId + '_TopSpan']; // [relative is the name of the "people picker div"]
		picker.DeleteProcessedUser();
	}
};
