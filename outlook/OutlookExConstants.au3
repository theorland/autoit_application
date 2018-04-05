#include-once

; #INDEX# =======================================================================================================================
; Title .........: OutlookEXConstants
; AutoIt Version : 3.3.6.1
; Language ......: English
; Description ...: Constants to be included in an AutoIt script when using the OutlookEX UDF.
; Author(s) .....: water
; Modified.......: 20120520 (YYYYMMDD)
; Contributors ..:
; Resources .....: Outlook 2007 Developer Reference -> Outlook Object Model Reference -> Enumerations: https://msdn.microsoft.com/en-us/library/bb259481(v=office.12).aspx
;                  Outlook 2010 Developer Reference -> Outlook Object Model Reference -> Enumerations: https://msdn.microsoft.com/en-us/library/ff860961(v=office.14).aspx
;                  Outlook 2013 Developer Reference -> Object Model -> Enumerations: https://msdn.microsoft.com/EN-US/library/ff860961.aspx
; ===============================================================================================================================

; #CONSTANTS# ===================================================================================================================
; MAPI names. Hex values for the Exchange properties that govern aging / archiving
; See: http://technet.microsoft.com/en-us/query/cc815492
Global Const $iPR_AGING_AGE_FOLDER = 0x6857000B ; BOOL Enable aging aka Archive for this folder: True = Enabled False = Disabled
Global Const $iPR_AGING_GRANULARITY = 0x36EE0003 ; LONG Aging granularity: 0 = Months 1 = Weeks 2 = Days
Global Const $iPR_AGING_PERIOD = 0x36EC0003 ; LONG, duration from 1 to 999 (combined with AGING GRANULARITY)
Global Const $iPR_AGING_DELETE_ITEMS = 0x6855000B ; BOOL FALSE = archive, TRUE = permanently delete
Global Const $iPR_AGING_FILE_NAME_AFTER9 = 0x6859001E ; STRING Path and filename of archive file for Exchange version > Exchange 9
Global Const $iPR_AGING_DEFAULT = 0x685E0003 ; LONG values unclear, seems like 3=do not archive, 1=archive according to defaults, 0=custom settings

; MsoAppLanguageID Enumeration. Specifies a language setting in a Microsoft Office application
; See: http://msdn.microsoft.com/en-us/library/aa432459%28v=office.12%29.aspx
Global Const $msoLanguageIDExeMode = 4 ; Execution mode language
Global Const $msoLanguageIDHelp = 3 ; Help language
Global Const $msoLanguageIDInstall = 1 ; Install language
Global Const $msoLanguageIDUI = 2 ; User interface language
Global Const $msoLanguageIDUIPrevious = 5 ; User interface language used prior to the current user interface language

; OlAddressEntryUserType Enumeration. Represents the type of user for the AddressEntry or object derived from AddressEntry.
; See http://msdn.microsoft.com/en-us/library/bb208051%28v=office.12%29.aspx
Global Const $olExchangeAgentAddressEntry = 3 ; An address entry that is an Exchange agent
Global Const $olExchangeDistributionListAddressEntry = 1 ; An address entry that is an Exchange distribution list
Global Const $olExchangeOrganizationAddressEntry = 4 ; An address entry that is an Exchange organization
Global Const $olExchangePublicFolderAddressEntry = 2 ; An address entry that is an Exchange public folder
Global Const $olExchangeRemoteUserAddressEntry = 5 ; An Exchange user that belongs to a different Exchange forest
Global Const $olExchangeUserAddressEntry = 0 ; An Exchange user that belongs to the same Exchange forest
Global Const $olLdapAddressEntry = 20 ; An address entry that uses the Lightweight Directory Access Protocol (LDAP)
Global Const $olOtherAddressEntry = 40 ; A custom or some other type of address entry such as FAX
Global Const $olOutlookContactAddressEntry = 10 ; An address entry in an Outlook Contacts folder
Global Const $olOutlookDistributionListAddressEntry = 11 ; An address entry that is an Outlook distribution list
Global Const $olSmtpAddressEntry = 30 ; An address entry that uses the Simple Mail Transfer Protocol (SMTP)

; OlAddressListType Enumeration. Represents the type of AddressList.
; See: http://msdn.microsoft.com/en-us/library/bb208052(v=office.12).aspx
Global Const $olCustomAddressList = 4 ; A custom address book provider
Global Const $olExchangeContainer = 1 ; A container for address lists on an Exchange server
Global Const $olExchangeGlobalAddressList = 0 ; An Exchange Global Address List
Global Const $olOutlookAddressList = 2 ; An address list that corresponds to the Outlook Contacts Address Book
Global Const $olOutlookLdapAddressList = 3 ; An address list that uses the Lightweight Directory Access Protocol (LDAP)

; OlAttachmentType Enumeration. Specifies the attachment type.
; See: http://msdn.microsoft.com/en-us/library/bb208055(v=office.12).aspx
Global Const $olByReference = 4 ; The attachment is a shortcut to the location of the original file
Global Const $olByValue = 1 ; The attachment is a copy of the original file and can be accessed even if the original file is removed
Global Const $olEmbeddeditem = 5 ; The attachment is an Outlook message format file (.msg) and is a copy of the original message
Global Const $olOLE = 6 ; The attachment is an OLE document

; OlBodyFormat Enumeration. Specifies the format of the body text of an item.
; See: http://msdn.microsoft.com/en-us/library/bb208058(v=office.12).aspx
Global Const $olFormatHTML = 2 ; HTML format
Global Const $olFormatPlain = 1 ; Plain format
Global Const $olFormatRichText = 3 ; Rich text format
Global Const $olFormatUnspecified = 0 ; Unspecified format

; OlBusyStatus Enumeration. Indicates a user's availability.
; See: http://msdn.microsoft.com/en-us/library/bb208060(v=office.12).aspx
Global Const $olBusy = 2 ; The user is busy
Global Const $olFree = 0 ; The user is available
Global Const $olOutOfOffice = 3 ; The user is out of office
Global Const $olTentative = 1 ; The user has a tentative appointment scheduled

; OlCategoryColor Enumeration. Indicates the color specified in the Color property of the Category object.
; See: http://msdn.microsoft.com/en-us/library/bb208064%28v=office.12%29.aspx
Global Const $olCategoryColorBlack = 15 ; Black
Global Const $olCategoryColorBlue = 8 ; Blue
Global Const $olCategoryColorDarkBlue = 23 ; Dark Blue
Global Const $olCategoryColorDarkGray = 14 ; Dark Gray
Global Const $olCategoryColorDarkGreen = 20 ; Dark Green
Global Const $olCategoryColorDarkMaroon = 25 ; Dark Maroon
Global Const $olCategoryColorDarkOlive = 22 ; Dark Olive
Global Const $olCategoryColorDarkOrange = 17 ; Dark Orange
Global Const $olCategoryColorDarkPeach = 18 ; Dark Peach
Global Const $olCategoryColorDarkPurple = 24 ; Dark Purple
Global Const $olCategoryColorDarkRed = 16 ; Dark Red
Global Const $olCategoryColorDarkSteel = 12 ; Dark Steel
Global Const $olCategoryColorDarkTeal = 21 ; Dark Teal
Global Const $olCategoryColorDarkYellow = 19 ; Dark Yellow
Global Const $olCategoryColorGray = 13 ; Gray
Global Const $olCategoryColorGreen = 5 ; Green
Global Const $olCategoryColorMaroon = 10 ; Maroon
Global Const $olCategoryColorNone = -1 ; No color assigned
Global Const $olCategoryColorOlive = 7 ; Olive
Global Const $olCategoryColorOrange = 2 ; Orange
Global Const $olCategoryColorPeach = 3 ; Peach
Global Const $olCategoryColorPurple = 9 ; Purple
Global Const $olCategoryColorRed = 1 ; Red
Global Const $olCategoryColorSteel = 11 ; Steel
Global Const $olCategoryColorTeal = 6 ; Teal
Global Const $olCategoryColorYellow = 4 ; Yellow

; OlCategoryShortcutKey Enumeration. Identifies the shortcut key specified for the ShortcutKey property of the Category object.
; See: http://msdn.microsoft.com/en-us/library/bb208065(v=office.12).aspx
Global Const $olCategoryShortcutKeyCtrlF10 = 10 ; CTRL+F10
Global Const $olCategoryShortcutKeyCtrlF11 = 11 ; CTRL+F11
Global Const $olCategoryShortcutKeyCtrlF12 = 12 ; CTRL+F12
Global Const $olCategoryShortcutKeyCtrlF2 = 2 ; CTRL+F2
Global Const $olCategoryShortcutKeyCtrlF3 = 3 ; CTRL+F3
Global Const $olCategoryShortcutKeyCtrlF4 = 4 ; CTRL+F4
Global Const $olCategoryShortcutKeyCtrlF5 = 5 ; CTRL+F5
Global Const $olCategoryShortcutKeyCtrlF6 = 6 ; CTRL+F6
Global Const $olCategoryShortcutKeyCtrlF7 = 7 ; CTRL+F7
Global Const $olCategoryShortcutKeyCtrlF8 = 8 ; CTRL+F8
Global Const $olCategoryShortcutKeyCtrlF9 = 9 ; CTRL+F9
Global Const $olCategoryShortcutKeyNone = 0 ; No shortcut key specified

; OlDaysOfWeek Enumeration. Constants representing days of the week.
; See: http://msdn.microsoft.com/en-us/library/bb208069(v=office.12).aspx
Global Const $olFriday = 32 ; Friday
Global Const $olMonday = 2 ; Monday
Global Const $olSaturday = 64 ; Saturday
Global Const $olSunday = 1 ; Sunday
Global Const $olThursday = 16 ; Thursday
Global Const $olTuesday = 4 ; Tuesday
Global Const $olWednesday = 8 ; Wednesday

; OlDefaultFolders Enumeration. Specifies the folder type for the current Outlook profile.
; See: http://msdn.microsoft.com/en-us/library/bb208072.aspx
Global Const $olFolderCalendar = 9 ; Calendar folder
Global Const $olFolderContacts = 10 ; The Contacts folder
Global Const $olFolderDeletedItems = 3 ; The Deleted Items folder
Global Const $olFolderDrafts = 16 ; The Drafts folder
Global Const $olFolderInbox = 6 ; The Inbox folder
Global Const $olFolderJournal = 11 ; The Journal folder
Global Const $olFolderJunk = 23 ; The Junk E-Mail folder
Global Const $olFolderNotes = 12 ; The Notes folder
Global Const $olFolderOutbox = 4 ; The Outbox folder
Global Const $olFolderSentMail = 5 ; The Sent Mail folder
Global Const $olFolderTasks = 13 ; The Tasks folder
Global Const $olFolderToDo = 28 ; The To Do folder
Global Const $olFolderRssFeeds = 25 ; The RSS Feeds folder

; OlDefaultSelectNamesDisplayMode Enumeration. Specifies the default caption, the number of buttons, the button labels, and the address lists to display in the Select Names dialog box.
; See: http://msdn.microsoft.com/en-us/library/bb208073%28v=office.12%29.aspx (for details about the set edit boxes, labels, buttons, caption and other settings)
Global Const $olDefaultDelegates = 6
Global Const $olDefaultMail = 1
Global Const $olDefaultMeeting = 2
Global Const $olDefaultMembers = 5
Global Const $olDefaultPickRooms = 8
Global Const $olDefaultSharingRequest = 4
Global Const $olDefaultSingleName = 7
Global Const $olDefaultTask = 3

; OlExchangeStoreType Enumeration. Specifies the type of Exchange store.
; See (Outlook 2013): https://msdn.microsoft.com/EN-US/library/ff864239.aspx
Global Const $olExchangeMailbox = 1 ; Specifies an Exchange delegate mailbox store
Global Const $olExchangePublicFolder = 2 ; Specifies an Exchange Public Folder store
Global Const $olNotExchange = 3 ; Specifies that the store is not an Exchange store
Global Const $olAdditionalExchangeMailbox = 4 ; Specifies an additional Exchange mailbox store
Global Const $olPrimaryExchangeMailbox = 0 ; Specifies a primary Exchange mailbox store

; OlImportance Enumeration. Specifies the level of importance for an item marked by the creator of the item.
; See: http://msdn.microsoft.com/en-us/library/bb208101(v=office.12).aspx
Global Const $olImportanceHigh = 2 ; Item is marked as high importance
Global Const $olImportanceLow = 0 ; Item is marked as low importance
Global Const $olImportanceNormal = 1 ; Item is marked as medium importance

; OlInspectorClose Enumeration. Indicates the save mode.
; See: http://msdn.microsoft.com/en-us/library/bb208102%28v=office.12%29.aspx
Global Const $olDiscard = 1 ; Changes to the document are discarded
Global Const $olPromptForSave = 2 ; User is prompted to save documents
Global Const $olSave = 0 ; Documents are saved

; OlItemType Enumeration. Indicates the Outlook Item type.
; See: http://msdn.microsoft.com/en-us/library/bb208104.aspx
Global Const $olAppointmentItem = 1 ; Represents an AppointmentItem
Global Const $olContactItem = 2 ; Represents a ContactItem
Global Const $olDistributionListItem = 7 ; Represents an DistListItem
Global Const $olJournalItem = 4 ; Represents a JournalItem
Global Const $olMailItem = 0 ; Represents a MailItem
Global Const $olNoteItem = 5 ; Represents a NoteItem
Global Const $olPostItem = 6 ; Represents a PostItem
Global Const $olTaskItem = 3 ; Represents a TaskItem

; OlMailRecipientType Enumeration. Indicates the recipient type for the Item.
; See: http://msdn.microsoft.com/en-us/library/bb208107(v=office.12).aspx
Global Const $olBCC = 3 ; The recipient is specified in the BCC property of the Item
Global Const $olCC = 2 ; The recipient is specified in the CC property of the Item
Global Const $olOriginator = 0 ; Originator (sender) of the Item
Global Const $olTo = 1 ; The recipient is specified in the To property of the Item

; OlMarkInterval Enumeration. Specifies the time period for which an Outlook item is marked as a task.
; See: http://msdn.microsoft.com/en-us/library/bb208108(v=office.12).aspx
Global Const $olMarkNoDate = 4 ; Mark the task due with no date
Global Const $olMarkNextWeek = 3 ; Mark the task due next week
Global Const $olMarkThisWeek = 2 ; Mark the task due this week
Global Const $olMarkToday = 0 ; Mark the task due today
Global Const $olMarkTomorrow = 1 ; Mark the task due tomorrow

; OlMeetingRecipientType Enumeration. Indicates the recipient type for the meeting.
; See: http://msdn.microsoft.com/en-us/library/bb208110(v=office.12).aspx
Global Const $olOptional = 2 ; Optional attendee
Global Const $olOrganizer = 0 ; Meeting organizer
Global Const $olRequired = 1 ; Required attendee
Global Const $olResource = 3 ; A resource such as a conference room

; OlMeetingResponse Enumeration. Indicates the response to a meeting request.
; See: http://msdn.microsoft.com/en-us/library/bb208112(v=office.12).aspx
Global Const $olMeetingAccepted = 3 ; The meeting was accepted
Global Const $olMeetingDeclined = 4 ; The meeting was declined
Global Const $olMeetingTentative = 2 ; The meeting was tentatively accepted

; OlMeetingStatus Enumeration. Indicates the status of the meeting.
; See: http://msdn.microsoft.com/en-us/library/bb208113(v=office.12).aspx
Global Const $olMeeting = 1 ; The meeting has been scheduled
Global Const $olMeetingCanceled = 5 ; The scheduled meeting has been cancelled
Global Const $olMeetingReceived = 3 ; The meeting request has been received
Global Const $olMeetingReceivedAndCanceled = 7 ; The scheduled meeting has been cancelled but still appears on the user's calendar
Global Const $olNonMeeting = 0 ; An Appointment item without attendees has been scheduled. This status can be used to set up holidays on a calendar

; OlNavigationModuleType Enumeration. Identifies the navigation module type of a NavigationModule object.
; See: http://msdn.microsoft.com/en-us/library/office/bb208114(v=office.12).aspx
Global Const $olModuleCalendar = 1 ; Identifies a CalendarModule object that represents the Calendar navigation module
Global Const $olModuleContacts = 2 ; Identifies a ContactsModule object that represents the Contacts navigation module
Global Const $olModuleFolderList = 6 ; Identifies a NavigationModule object that represents the Folders List navigation module
Global Const $olModuleJournal = 4 ; Identifies a JournalModule object that represents the Journal navigation module
Global Const $olModuleMail = 0 ; Identifies a MailModule object that represents the Mail navigation module
Global Const $olModuleNotes = 5 ; Identifies a NotesModule object that represents the Notes navigation module
Global Const $olModuleShortcuts = 7 ; Identifies a NavigationModule object that represents the Shortcuts navigation module
Global Const $olModuleTasks = 3 ; Identifies a TasksModule object that represents the Tasks navigation module

; OlObjectClass Enumeration. Constants representing the different Microsoft Office Outlook object classes.
; See: http://msdn.microsoft.com/en-us/library/bb208118(v=office.12).aspx
Global Const $olAccount = 105 ; Represents an Account object
Global Const $olAccountRuleCondition = 135 ; Represents an AccountRuleCondition object
Global Const $olAccounts = 106 ; Represents an Accounts object
Global Const $olAction = 32 ; Represents an Action object
Global Const $olActions = 33 ; Represents an Actions object
Global Const $olAddressEntries = 21 ; Represents an AddressEntries object
Global Const $olAddressEntry = 8 ; Represents an AddressEntry object
Global Const $olAddressList = 7 ; Represents an AddressList object
Global Const $olAddressLists = 20 ; Represents an AddressLists object
Global Const $olAddressRuleCondition = 170 ; Represents an AddressRuleCondition object
Global Const $olApplication = 0 ; Represents an Application object
Global Const $olAppointment = 26 ; Represents an AppointmentItem object
Global Const $olAssignToCategoryRuleAction = 122 ; Represents an AssignToCategoryRuleAction object
Global Const $olAttachment = 5 ; Represents an Attachment object
Global Const $olAttachments = 18 ; Represents an Attachments object
Global Const $olAttachmentSelection = 169 ; Represents an AttachmentSelection object
Global Const $olAutoFormatRule = 147 ; Represents an AutoFormatRule object
Global Const $olAutoFormatRules = 148 ; Represents an AutoFormatRules object
Global Const $olCalendarModule = 159 ; Represents a CalendarModule object
Global Const $olCalendarSharing = 151 ; Represents a CalendarSharing object
Global Const $olCategories = 153 ; Represents a Categories object
Global Const $olCategory = 152 ; Represents a Category object
Global Const $olCategoryRuleCondition = 130 ; Represents a CategoryRuleCondition object
Global Const $olClassBusinessCardView = 168 ; Represents a BusinessCardView object
Global Const $olClassCalendarView = 139 ; Represents a CalendarView object
Global Const $olClassCardView = 138 ; Represents a CardView object
Global Const $olClassIconView = 137 ; Represents a IconView object
Global Const $olClassNavigationPane = 155 ; Represents a NavigationPane object
Global Const $olClassTableView = 136 ; Represents a TableView object
Global Const $olClassTimeLineView = 140 ; Represents a TimelineView object
Global Const $olClassTimeZone = 174 ; Represents a TimeZone object
Global Const $olClassTimeZones = 175 ; Represents a TimeZones object
Global Const $olColumn = 154 ; Represents a Column object
Global Const $olColumnFormat = 149 ; Represents a ColumnFormat object
Global Const $olColumns = 150 ; Represents a Columns object
Global Const $olConflict = 102 ; Represents a Conflict object
Global Const $olConflicts = 103 ; Represents a Conflicts object
Global Const $olContact = 40 ; Represents a ContactItem object
Global Const $olContactsModule = 160 ; Represents a ContactsModule object
Global Const $olDistributionList = 69 ; Represents a ExchangeDistributionList object
Global Const $olDocument = 41 ; Represents a DocumentItem object
Global Const $olException = 30 ; Represents an Exception object
Global Const $olExceptions = 29 ; Represents an Exceptions object
Global Const $olExchangeDistributionList = 111 ; Represents an ExchangeDistributionList object
Global Const $olExchangeUser = 110 ; Represents an ExchangeUser object
Global Const $olExplorer = 34 ; Represents an Explorer object
Global Const $olExplorers = 60 ; Represents an Explorers object
Global Const $olFolder = 2 ; Represents a Folder object
Global Const $olFolders = 15 ; Represents a Folders object
Global Const $olFolderUserProperties = 172 ; Represents a UserDefinedProperties object
Global Const $olFolderUserProperty = 171 ; Represents a UserDefinedProperty object
Global Const $olFormDescription = 37 ; Represents a FormDescription object
Global Const $olFormNameRuleCondition = 131 ; Represents a FormNameRuleCondition object
Global Const $olFormRegion = 129 ; Represents a FormRegion object
Global Const $olFromRssFeedRuleCondition = 173 ; Represents a FromRssFeedRuleCondition object
Global Const $olFromRuleCondition = 132 ; Represents a ToOrFromRuleCondition object
Global Const $olImportanceRuleCondition = 128 ; Represents an ImportanceRuleCondition object
Global Const $olInspector = 35 ; Represents an Inspector object
Global Const $olInspectors = 61 ; Represents an Inspectors object
Global Const $olItemProperties = 98 ; Represents an ItemProperties object
Global Const $olItemProperty = 99 ; Represents an ItemProperty object
Global Const $olItems = 16 ; Represents an Items object
Global Const $olJournal = 42 ; Represents a JournalItem object
Global Const $olJournalModule = 162 ; Represents a JournalModule object
Global Const $olLink = 75 ; Represents a Link object
Global Const $olLinks = 76 ; Represents a Links object
Global Const $olMail = 43 ; Represents a MailItem object
Global Const $olMailModule = 158 ; Represents a MailModule object
Global Const $olMarkAsTaskRuleAction = 124 ; Represents a MarkAsTaskRuleAction object
Global Const $olMeetingCancellation = 54 ; Represents a MeetingItem object that is a meeting cancellation notice
Global Const $olMeetingRequest = 53 ; Represents a MeetingItem object that is a meeting request
Global Const $olMeetingResponseNegative = 55 ; Represents a MeetingItem object that is a refusal of a meeting request
Global Const $olMeetingResponsePositive = 56 ; Represents a MeetingItem object that is an acceptance of a meeting request
Global Const $olMeetingResponseTentative = 57 ; Represents a MeetingItem object that is a tentative acceptance of a meeting request
Global Const $olMoveOrCopyRuleAction = 118 ; Represents a MoveOrCopyRuleAction object
Global Const $olNamespace = 1 ; Represents a NameSpace object
Global Const $olNavigationFolder = 167 ; Represents a NavigationFolder object
Global Const $olNavigationFolders = 166 ; Represents a NavigationFolders object
Global Const $olNavigationGroup = 165 ; Represents a NavigationGroup object
Global Const $olNavigationGroups = 164 ; Represents a NavigationGroups object
Global Const $olNavigationModule = 157 ; Represents a NavigationModule object
Global Const $olNavigationModules = 156 ; Represents a NavigationModules object
Global Const $olNewItemAlertRuleAction = 125 ; Represents a NewItemAlertRuleAction object
Global Const $olNote = 44 ; Represents a NoteItem object
Global Const $olNotesModule = 163 ; Represents a NotesModule object
Global Const $olOrderField = 144 ; Represents an OrderField object
Global Const $olOrderFields = 145 ; Represents an OrderFields object
Global Const $olOutlookBarGroup = 66 ; Represents an OutlookBarGroup object
Global Const $olOutlookBarGroups = 65 ; Represents an OutlookBarGroups object
Global Const $olOutlookBarPane = 63 ; Represents an OutlookBarPane object
Global Const $olOutlookBarShortcut = 68 ; Represents an OutlookBarShortcut object
Global Const $olOutlookBarShortcuts = 67 ; Represents an OutlookBarShortcuts object
Global Const $olOutlookBarStorage = 64 ; Represents an OutlookBarStorage object
Global Const $olPages = 36 ; Represents a Pages object
Global Const $olPanes = 62 ; Represents a Panes object
Global Const $olPlaySoundRuleAction = 123 ; Represents a PlaySoundRuleAction object
Global Const $olPost = 45 ; Represents a PostItem object
Global Const $olPropertyAccessor = 112 ; Represents a PropertyAccessor object
Global Const $olPropertyPages = 71 ; Represents a PropertyPages object
Global Const $olPropertyPageSite = 70 ; Represents a PropertyPageSite object
Global Const $olRecipient = 4 ; Represents a Recipient object
Global Const $olRecipients = 17 ; Represents a Recipients object
Global Const $olRecurrencePattern = 28 ; Represents a RecurrencePattern object
Global Const $olReminder = 101 ; Represents a Reminder object
Global Const $olReminders = 100 ; Represents a Reminders object
Global Const $olRemote = 47 ; Represents a RemoteItem object
Global Const $olReport = 46 ; Represents a ReportItem object
Global Const $olResults = 78 ; Represents a Results object
Global Const $olRow = 121 ; Represents a Row object
Global Const $olRule = 115 ; Represents a Rule object
Global Const $olRuleAction = 117 ; Represents a RuleAction object
Global Const $olRuleActions = 116 ; Represents a RuleAction object
Global Const $olRuleCondition = 127 ; Represents a RuleCondition object
Global Const $olRuleConditions = 126 ; Represents a RuleConditions object
Global Const $olRules = 114 ; Represents a Rules object
Global Const $olSearch = 77 ; Represents a Search object
Global Const $olSelection = 74 ; Represents a Selection object
Global Const $olSelectNamesDialog = 109 ; Represents a SelectNamesDialog object
Global Const $olSenderInAddressListRuleCondition = 133 ; Represents a SenderInAddressListRuleCondition object
Global Const $olSendRuleAction = 119 ; Represents a SendRuleAction object
Global Const $olSharing = 104 ; Represents a SharingItem object
Global Const $olStorageItem = 113 ; Represents a StorageItem object
Global Const $olStore = 107 ; Represents a Store object
Global Const $olStores = 108 ; Represents a Stores object
Global Const $olSyncObject = 72 ; Represents a SyncObject object
Global Const $olSyncObjects = 73 ; Represents a SyncObject object
Global Const $olTable = 120 ; Represents a Table object
Global Const $olTask = 48 ; Represents a TaskItem object
Global Const $olTaskRequest = 49 ; Represents a TaskRequestItem object
Global Const $olTaskRequestAccept = 51 ; Represents a TaskRequestAcceptItem object
Global Const $olTaskRequestDecline = 52 ; Represents a TaskRequestDeclineItem object
Global Const $olTaskRequestUpdate = 50 ; Represents a TaskRequestUpdateItem object
Global Const $olTasksModule = 161 ; Represents a TasksModule object
Global Const $olTextRuleCondition = 134 ; Represents a TextRuleCondition object
Global Const $olUserDefinedProperties = 172 ; Represents a UserDefinedProperties object
Global Const $olUserDefinedProperty = 171 ; Represents a UserDefinedProperty object
Global Const $olUserProperties = 38 ; Represents a UserProperties object
Global Const $olUserProperty = 39 ; Represents a UserProperty object
Global Const $olView = 80 ; Represents a View object
Global Const $olViewField = 142 ; Represents a ViewField object
Global Const $olViewFields = 141 ; Represents a ViewFields object
Global Const $olViewFont = 146 ; Represents a ViewFont object
Global Const $olViews = 79 ; Represents a Views object

; OlPane Enumeration. Specifies the pane (window).
; See: http://msdn.microsoft.com/en-us/library/bb208119(v=office.12).aspx
Global Const $olFolderList = 2 ; Folder list pane
Global Const $olNavigationPane = 4 ; Navigation pane
Global Const $olOutlookBar = 1 ; Outlook bar pane
Global Const $olPreview = 3 ; Preview pane
Global Const $olToDoBar = 5 ; To Do bar pane

; OlRecurrenceType Enumeration. Specifies the recurrence pattern type.
; See: http://msdn.microsoft.com/en-us/library/bb208124(v=office.12).aspx
Global Const $olRecursDaily = 0 ; Represents a daily recurrence pattern
Global Const $olRecursMonthly = 2 ; Represents a monthly recurrence pattern
Global Const $olRecursMonthNth = 3 ; Represents a MonthNth recurrence pattern. See RecurrencePattern.Instance property
Global Const $olRecursWeekly = 1 ; Represents a weekly recurrence pattern
Global Const $olRecursYearly = 5 ; Represents a yearly recurrence pattern
Global Const $olRecursYearNth = 6 ; Represents a YearNth recurrence pattern. See RecurrencePattern.Instance property

; OlRuleActionType Enumeration. Specifies the type of rule action for a rule.
; See: http://msdn.microsoft.com/en-us/library/bb208128%28v=office.12%29.aspx
Global Const $olRuleActionAssignToCategory = 2 ; Rule action is to assign categories to the message
Global Const $olRuleActionCcMessage = 27 ; Rule action is to cc the message to specified recipients
Global Const $olRuleActionClearCategories = 30 ; Rule action is to clear all the categories assigned to the message
Global Const $olRuleActionCopyToFolder = 5 ; Rule action is to copy the message to a specified folder
Global Const $olRuleActionCustomAction = 22 ; Rule action is to perform a custom action
Global Const $olRuleActionDefer = 28 ; Rule action is to defer delivery of the message by the specified number of minutes
Global Const $olRuleActionDelete = 3 ; Rule action is to delete the message
Global Const $olRuleActionDeletePermanently = 4 ; Rule action is to permanently delete the message
Global Const $olRuleActionDesktopAlert = 24 ; Rule action is to display a desktop alert
Global Const $olRuleActionFlagClear = 13 ; Rule action is to clear the message flag
Global Const $olRuleActionFlagColor = 12 ; Rule action is to flag the message with a specified colored flag
Global Const $olRuleActionFlagForActionInDays = 11 ; Rule action is to flag the message for action in the specified number of days
Global Const $olRuleActionForward = 6 ; Rule action is to forward the message to the specified recipients
Global Const $olRuleActionForwardAsAttachment = 7 ; Rule action is to forward the message as an attachment to the specified recipients
Global Const $olRuleActionImportance = 14 ; Rule action is to mark the message with the specified level of importance
Global Const $olRuleActionMarkAsTask = 29 ; Rule action is to mark the message as a task (Bug in the docu: docu says 41. Wrong hex to dec conversion)
Global Const $olRuleActionMarkRead = 19 ; Rule action is to mark the message as read
Global Const $olRuleActionMoveToFolder = 1 ; Rule action is to move the message to the specified folder
Global Const $olRuleActionNewItemAlert = 23 ; Rule action is to display the specified text in the New Item Alert dialog box
Global Const $olRuleActionNotifyDelivery = 26 ; Rule action is to request delivery notification for the message being sent
Global Const $olRuleActionNotifyRead = 25 ; Rule action is to request read notification for the message being sent
Global Const $olRuleActionPlaySound = 17 ; Rule action is to play a sound file
Global Const $olRuleActionPrint = 16 ; Rule action is to print the message on the default printer
Global Const $olRuleActionRedirect = 8 ; Rule action is to redirect the message to the specified recipients
Global Const $olRuleActionRunScript = 20 ; Rule action is to run a script
Global Const $olRuleActionSensitivity = 15 ; Rule action is to mark the message with the specified level of sensitivity
Global Const $olRuleActionServerReply = 9 ; Rule action is to request the server to reply with the specified mail item
Global Const $olRuleActionStartApplication = 18 ; Rule action is to run an .exe file
Global Const $olRuleActionStop = 21 ; Rule action is to stop processing more rules
Global Const $olRuleActionTemplate = 10 ; Rule action is to use the specified template (.oft) file as a form template
Global Const $olRuleActionUnknown = 0 ; Unrecognized rule action

; OlRuleConditionType Enumeration. Specifies the type of rule condition or exception condition of a rule.
; See: http://msdn.microsoft.com/en-us/library/bb208129%28v=office.12%29.aspx
Global Const $olConditionAccount = 3 ; Account is the account specified in AccountRuleCondition.Account
Global Const $olConditionAnyCategory = 29 ; Message is assigned to any category
Global Const $olConditionBody = 13 ; Body contains words specified in TextRuleCondition.Text
Global Const $olConditionBodyOrSubject = 14 ; Body or subject contains words specified by TextRuleCondition.Text
Global Const $olConditionCategory = 18 ; Category is the category specified in CategoryRuleCondition.Categories
Global Const $olConditionCc = 9 ; Message has my name in the Cc box
Global Const $olConditionDateRange = 22 ; Message was received between x and y, where x and y are Date values
Global Const $olConditionFlaggedForAction = 8 ; Message is flagged for the specified action
Global Const $olConditionFormName = 23 ; Message uses the form specified in FormNameRuleCondition.FormName
Global Const $olConditionFrom = 1 ; Sender is in the recipient list specified in ToOrFromRuleCondition.Recipients
Global Const $olConditionFromAnyRssFeed = 31 ; Message is generated from any RSS subscription
Global Const $olConditionFromRssFeed = 30 ; Message is generated from a specific RSS subscription
Global Const $olConditionHasAttachment = 20 ; Message has one or more attachments
Global Const $olConditionImportance = 6 ; Message is marked with the specified level of importance
Global Const $olConditionLocalMachineOnly = 27 ; Rule can run only on the local machine
Global Const $olConditionMeetingInviteOrUpdate = 26 ; Message is a meeting invitation or update
Global Const $olConditionMessageHeader = 15 ; Message header contains words specified in TextRuleCondition.Text
Global Const $olConditionNotTo = 11 ; Message does not have my name in the To box
Global Const $olConditionOnlyToMe = 4 ; Message is sent only to me
Global Const $olConditionOOF = 19 ; Message is an out-of-office message
Global Const $olConditionOtherMachine = 28 ; Rule can run only on a specific machine that is not the current machine
Global Const $olConditionProperty = 24 ; Document property is exactly, contains, or does not contain specified properties
Global Const $olConditionRecipientAddress = 16 ; Recipient address contains words specified in TextRuleCondition.Text
Global Const $olConditionSenderAddress = 17 ; Sender address contains words specified in TextRuleCondition.Text
Global Const $olConditionSenderInAddressBook = 25 ; Sender is in the address list specified in AddressRuleCondition.Address
Global Const $olConditionSensitivity = 7 ; Message is marked with the specified level of sensitivity
Global Const $olConditionSentTo = 12 ; Sent to recipients (To, Cc) are in the recipient list specified in ToOrFromRuleCondition.Recipients
Global Const $olConditionSizeRange = 21 ; Message size is between x and y in units of KB, where x and y are Integer values
Global Const $olConditionSubject = 2 ; Subject contains words specified in TextRuleCondition.Text
Global Const $olConditionTo = 5 ; My name is in the To box
Global Const $olConditionToOrCc = 10 ; Message has my name in the To or Cc box
Global Const $olConditionUnknown = 0 ; Unrecognized condition

; OlRuleExecuteOption Enumeration. Specifies the type of messages in the specified folder or folders that a rule should be applied to.
; See: http://msdn.microsoft.com/en-us/library/bb208130(v=office.12).aspx
Global Const $olRuleExecuteAllMessages = 0 ; Executes a rule against all messages in the specified folder or folders
Global Const $olRuleExecuteReadMessages = 1 ; Executes a rule against messages that have been read in the specified folder or folders
Global Const $olRuleExecuteUnreadMessages = 2 ; Executes a rule against messages that have not been read in the specified folder or folders

; OlRuleType Enumeration. Indicates if a rule is applied to messages that are being sent or received.
; See: http://msdn.microsoft.com/en-us/library/bb208131(v=office.12).aspx
Global Const $olRuleReceive = 0 ; Indicates that the rule is applied to messages that are being received
Global Const $olRuleSend = 1 ; Indicates that the rule is being applied to messages being sent

; OlSaveAsType Enumeration. Specifies the file format.
; See: http://msdn.microsoft.com/en-us/library/bb208132(v=office.12).aspx
Global Const $olDoc = 4 ; Microsoft Office Word format (.doc)
Global Const $olHTML = 5 ; HTML format (.html)
Global Const $olICal = 8 ; iCal format (.ics)
Global Const $olMHTML = 10 ; MIME HTML format (.mht)
Global Const $olMSG = 3 ; Outlook message format (.msg)
Global Const $olMSGUnicode = 9 ; Outlook Unicode message format (.msg)
Global Const $olRTF = 1 ; Rich Text format (.rtf)
Global Const $olTemplate = 2 ; Microsoft Office Outlook template (.oft)
Global Const $olTXT = 0 ; Text format (.txt)
Global Const $olVCal = 7 ; VCal format (.vcs)
Global Const $olVCard = 6 ; VCard format (.vcf)

; OlSensitivity Enumeration. Represents the sensitivity for the Microsoft Office Outlook item.
; See: http://msdn.microsoft.com/en-us/library/bb208134(v=office.12).aspx
Global Const $olConfidential = 3 ; Confidential
Global Const $olNormal = 0 ; Normal sensitivity
Global Const $olPersonal = 1 ; Personal
Global Const $olPrivate = 2 ; Private

; OlShowItemCount Enumeration. Indicates which type of Microsoft Outlook item count is displayed for folders in the Outlook navigation pane.
; See: http://msdn.microsoft.com/en-us/library/bb208137(v=office.12).aspx
Global Const $olNoItemCount = 0 ; No item count displayed
Global Const $olShowTotalItemCount = 2 ; Shows count of total number of items
Global Const $olShowUnreadItemCount = 1 ; Shows count of unread items

; OlStorageIdentifierType Enumeration. Specifies the type of identifier for a StorageItem object.
; See: http://msdn.microsoft.com/en-us/library/bb208140(v=office.12).aspx
Global Const $olIdentifyByEntryID = 1 ; Identifies a StorageItem by EntryID
Global Const $olIdentifyByMessageClass = 2 ; Identifies a StorageItem by message class
Global Const $olIdentifyBySubject = 0 ; Identifies a StorageItem by Subject

; OlStoreType Enumeration. Indicates the format in which the data file should be created.
; See: http://msdn.microsoft.com/en-us/library/bb208141(v=office.12).aspx
Global Const $olStoreANSI = 3 ; ANSI format personal folders file (.pst) compatible with all previous versions of Microsoft Office Outlook format
Global Const $olStoreDefault = 1 ; Default format compatible with the mailbox mode in which Microsoft Office Outlook runs on the Microsoft Exchange Server
Global Const $olStoreUnicode = 2 ; Unicode format personal folders file (.pst) compatible with Microsoft Office Outlook 2003 and later

; OlTaskRecipientType Enumeration. Indicates the task recipient type.
; See: http://msdn.microsoft.com/en-us/library/bb208145(v=office.12).aspx
Global Const $olFinalStatus = 3 ; Indicates that the recipient will receive completion reports for the task
Global Const $olUpdate = 2 ; Indicates that the recipient will receive status updates for the task

; OlTaskResponse Enumeration. Indicates the response to a task request.
; See: http://msdn.microsoft.com/en-us/library/bb208146(v=office.12).aspx
Global Const $olTaskAccept = 2 ; Task accepted
Global Const $olTaskAssign = 1 ; Task reassigned
Global Const $olTaskDecline = 3 ; Task declined
Global Const $olTaskSimple = 0 ; Task is a simple task and cannot be accepted, declined, or assigned. Not a valid parameter to the TaskItem.Respond method

; OlUserPropertyType Enumeration. Indicates the user property type.
; See: http://msdn.microsoft.com/en-us/library/bb208152(v=office.12).aspx
Global Const $olCombination = 19 ; The property type is a combination of other types
Global Const $olCurrency = 14 ; Represents a Currency property type
Global Const $olDateTime = 5 ; Represents a DateTime property type
Global Const $olDuration = 7 ; Represents a time duration property type
Global Const $olEnumeration = 21 ; Represents an enumeration property type
Global Const $olFormula = 18 ; Represents a formula property type. See UserDefinedProperty.Formula property
Global Const $olInteger = 20 ; Represents an Integer number property type
Global Const $olKeywords = 11 ; Represents a String array property type used to store keywords
Global Const $olNumber = 3 ; Represents a Double number property type
Global Const $olOutlookInternal = 0 ; Represents an Outlook internal property type.
Global Const $olPercent = 12 ; Represents a Double number property type used to store a percentage
Global Const $olSmartFrom = 22 ; Represents a smart from property type
Global Const $olText = 1 ; Represents a String property type
Global Const $olYesNo = 6 ; Represents a yes/no (Boolean) property type

; OlWindowState Enumeration. Indicates the window state.
; See: http://msdn.microsoft.com/en-us/library/bb208155(v=office.12).aspx
Global Const $olMaximized = 0 ; The window is maximized
Global Const $olMinimized = 1 ; The window is minimized
Global Const $olNormalWindow = 2 ; The window is in the normal state (not minimized or maximized)
; ===============================================================================================================================
