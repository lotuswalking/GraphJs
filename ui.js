// Select DOM elements to work with
const authenticatedNav = document.getElementById("authenticated-nav");
const accountNav = document.getElementById("account-nav");
const mainContainer = document.getElementById("main-container");

const Views = { error: 1, home: 2, calendar: 3, mail: 4, presence: 5, mailRead: 6 };
// 自行封装的函数,创建bootstrap对象,动态显示class
function createElement(type, className, text) {
  var element = document.createElement(type);
  element.className = className;

  if (text) {
    var textNode = document.createTextNode(text);
    element.appendChild(textNode);
  }

  return element;
}

// 显示已认证适合需要显示的页面内容,属于home页面

function showAuthenticatedNav(account, view) {
  authenticatedNav.innerHTML = "";
  // 如果已经登录
  if (account) {
    //显示邮件按钮
    var emailNav = createElement("li", "nav-item");
    var emailLink = createElement(
      "button",
      `btn btn-link nav-link${view === Views.email ? "active" : ""}`,
      "Email"
    );
    emailLink.setAttribute("onclick", "getEmails();");
    emailNav.appendChild(emailLink);
    authenticatedNav.appendChild(emailNav);
    // Add Calendar link显示日历按钮
    var calendarNav = createElement("li", "nav-item");

    var calendarLink = createElement(
      "button",
      `btn btn-link nav-link${view === Views.calendar ? " active" : ""}`,
      "Calendar"
    );
    calendarLink.setAttribute("onclick", "getEvents();");
    calendarNav.appendChild(calendarLink);

    authenticatedNav.appendChild(calendarNav);
    //btn show presence
    var presence = createElement("li", "nav-item");
    var presenceLink = createElement(
      "button",
      `btn btn-lin nav-link${view === Views.presence ? "active" : ""}`,
      "Presence"
    );
    presenceLink.setAttribute("onclick", "getPresence()");
    presence.appendChild(presenceLink);
    authenticatedNav.appendChild(presence);
  }
}

function showAccountNav(account) {
  accountNav.innerHTML = "";
  // 如果已经认证,则显示用户信息
  if (account) {
    // Show the "signed-in" nav
    accountNav.className = "nav-item dropdown";
    var dropdown = createElement("a", "nav-link dropdown-toggle");
    dropdown.setAttribute("data-toggle", "dropdown");
    dropdown.setAttribute("role", "button");
    accountNav.appendChild(dropdown);

    var userIcon = createElement(
      "i",
      "far fa-user-circle fa-lg rounded-circle align-self-center"
    );
    userIcon.style.width = "32px";
    dropdown.appendChild(userIcon);

    var menu = createElement("div", "dropdown-menu dropdown-menu-right");
    dropdown.appendChild(menu);

    var userName = createElement("h5", "dropdown-item-text mb-0", account.name);
    menu.appendChild(userName);

    var userEmail = createElement(
      "p",
      "dropdown-item-text text-muted mb-0",
      account.userName
    );
    menu.appendChild(userEmail);

    var divider = createElement("div", "dropdown-divider");
    menu.appendChild(divider);

    var signOutButton = createElement("button", "dropdown-item", "Sign out");
    signOutButton.setAttribute("onclick", "signOut();");
    menu.appendChild(signOutButton);
  } else {
    // Show a "sign in" button
    accountNav.className = "nav-item";

    var signInButton = createElement(
      "button",
      "btn btn-link nav-link",
      "Sign in"
    );
    signInButton.setAttribute("onclick", "signIn();");
    accountNav.appendChild(signInButton);
  }
}

function showWelcomeMessage(account) {
  // Create jumbotron
  var jumbotron = createElement("div", "jumbotron");

  var heading = createElement("h1", null, "JavaScript SPA Graph Tutorial");
  jumbotron.appendChild(heading);

  var lead = createElement(
    "p",
    "lead",
    "This sample app shows how to use the Microsoft Graph API to access" +
      " a user's data from JavaScript."
  );
  jumbotron.appendChild(lead);

  if (account) {
    // Welcome the user by name
    var welcomeMessage = createElement("h4", null, `Welcome ${account.name}!`);
    // var welcomeMessage = createElement(
    //   "h4",
    //   null,
    //   `Your Email address  ${account.emailAddress}!`
    // );
    jumbotron.appendChild(welcomeMessage);

    var callToAction = createElement(
      "p",
      null,
      "Use the navigation bar at the top of the page to get started."
    );
    jumbotron.appendChild(callToAction);
  } else {
    // Show a sign in button in the jumbotron
    var signInButton = createElement(
      "button",
      "btn btn-primary btn-large",
      "Click here to sign in"
    );
    signInButton.setAttribute("onclick", "signIn();");
    jumbotron.appendChild(signInButton);
  }

  mainContainer.innerHTML = "";
  mainContainer.appendChild(jumbotron);
}
// 显示错误?
function showError(error) {
  var alert = createElement("div", "alert alert-danger");

  var message = createElement("p", "mb-3", error.message);
  alert.appendChild(message);

  if (error.debug) {
    var pre = createElement("pre", "alert-pre border bg-light p-2");
    alert.appendChild(pre);

    var code = createElement(
      "code",
      "text-break text-wrap",
      JSON.stringify(error.debug, null, 2)
    );
    pre.appendChild(code);
  }

  mainContainer.innerHTML = "";
  mainContainer.appendChild(alert);
}
// 显示主框架页面内容
function updatePage(account, view, data) {
  //    如果没有账号登录或者View的入口参数为空
  if (!view || !account) {
    view = Views.home;
  }
  //   显示右上角登录小下拉条
  showAccountNav(account);

  // 显示Email/Calendar/Presence等按钮
  showAuthenticatedNav(account, view);

  switch (view) {
    case Views.error:
      showError(data);
      break;
    case Views.home:
      showWelcomeMessage(account);
      break;
    case Views.calendar:
      showCalendar(data);
      break;
    case Views.mail:
      showEmail(data);
      break;
    case Views.presence:
      showPresence(data);
      break;
    case Views.mailRead:
      ShowEmailDetail(data);
      break;
  }
}

updatePage(null, Views.home);
// 显示在线信息
function showPresence(presence) {
//   console.log(presence);
  var div = document.createElement("div");
  div.appendChild(
    createElement("h1", null, "My Teams Status:" + presence.availability)
  );
  div.appendChild(createElement("h1", null, "activity:" + presence.activity));
  div.appendChild(createElement("h2", null, "ID: " + presence.id));
  mainContainer.innerHTML = "";
  mainContainer.appendChild(div);
}
// 显示邮件
function showEmail(emails) {
  try {
    // console.log(emails);
    var div = document.createElement("div");
    div.appendChild(createElement("h1", null, "Email"));
    var table = createElement("table", "table");
    div.appendChild(table);

    var thead = document.createElement("thead");
    table.appendChild(thead);

    var headerrow = document.createElement("tr");
    thead.appendChild(headerrow);

    var subject = createElement("th", null, "From");
    subject.setAttribute("scope", "col");
    headerrow.appendChild(subject);

    var subject = createElement("th", null, "Subject");
    subject.setAttribute("scope", "col");
    headerrow.appendChild(subject);

    var receivedDateTime = createElement("th", null, "receivedDateTime");
    receivedDateTime.setAttribute("scope", "col");
    headerrow.appendChild(receivedDateTime);
    var tbody = document.createElement("tbody");
    table.appendChild(tbody);

    for (const mail of emails.value) {
      var mailrow = document.createElement("tr");
      var mailrow = document.createElement("tr");
      mailrow.setAttribute("key", mail.id);
      mailrow.setAttribute('onclick','getEmailDetail("'+ mail.id +'")');
      tbody.appendChild(mailrow);

      var fromcell = createElement("td", null, mail.from.emailAddress.address);
      mailrow.appendChild(fromcell);
      var subjectCell = createElement("td", null, mail.subject);
    //   subjectCell.setAttribute('onclick','ShowEmailDetail(this)');

      mailrow.appendChild(subjectCell);
      var startcell = createElement(
        "td",
        null,
        moment
          .utc(mail.receivedDateTime.dateTime)
          .local()
          .format("M/D/YY h:mm A")
      );
      mailrow.appendChild(startcell);
    }

    mainContainer.innerHTML = "";
    mainContainer.appendChild(div);
  } catch (ex) {
    console.log(ex);
  }
}

function ShowEmailDetail(emaildata)
{
    // console.log(emaildata);
    var div = document.createElement("div");
    div.appendChild(createElement("p", null,'Subject:' + emaildata.subject));
    div.appendChild(createElement("p", null,'from:' + emaildata.from.emailAddress.name +' ' +emaildata.from.emailAddress.address));
    var body = document.createElement('div');
    body.innerHTML = emaildata.body.content;
    div.appendChild(body);
    mainContainer.innerHTML = '';
    mainContainer.appendChild(div);
    //  alert('item.subject'+src);

}

// 显示日历
function showCalendar(events) {
  var div = document.createElement("div");

  div.appendChild(createElement("h1", null, "Calendar"));

  var table = createElement("table", "table");
  div.appendChild(table);

  var thead = document.createElement("thead");
  table.appendChild(thead);

  var headerrow = document.createElement("tr");
  thead.appendChild(headerrow);

  var organizer = createElement("th", null, "Organizer");
  organizer.setAttribute("scope", "col");
  headerrow.appendChild(organizer);

  var subject = createElement("th", null, "Subject");
  subject.setAttribute("scope", "col");
  headerrow.appendChild(subject);

  var start = createElement("th", null, "Start");
  start.setAttribute("scope", "col");
  headerrow.appendChild(start);

  var end = createElement("th", null, "End");
  end.setAttribute("scope", "col");
  headerrow.appendChild(end);

  var tbody = document.createElement("tbody");
  table.appendChild(tbody);

  for (const event of events.value) {
    var eventrow = document.createElement("tr");
    eventrow.setAttribute("key", event.id);
    tbody.appendChild(eventrow);

    var organizercell = createElement(
      "td",
      null,
      event.organizer.emailAddress.name
    );
    eventrow.appendChild(organizercell);

    var subjectcell = createElement("td", null, event.subject);
    eventrow.appendChild(subjectcell);

    var startcell = createElement(
      "td",
      null,
      moment.utc(event.start.dateTime).local().format("M/D/YY h:mm A")
    );
    eventrow.appendChild(startcell);

    var endcell = createElement(
      "td",
      null,
      moment.utc(event.end.dateTime).local().format("M/D/YY h:mm A")
    );
    eventrow.appendChild(endcell);
  }

  mainContainer.innerHTML = "";
  mainContainer.appendChild(div);
}
