
function Timesheet(timesheet) {

  var _timesheet = timesheet ? timesheet : [];

  this.add = function(date, client, duration) {
    if (!hasDayOfMonth(date.getDate())) {
      create(date);
    }
    var timesheetDay = get(date.getDate());
    timesheetDay.add(client, duration);
  };

  this.timesheet = function() {
    return _timesheet;
  };

  this.iterator = function() {
    var nextIndex = 0;

    return {
      next: function() {
        while (nextIndex < _timesheet.length) {
          var timesheetDay = _timesheet[nextIndex++];
          if (timesheetDay) {
            return {value: timesheetDay, done: false};
          }
        }
        if (_timesheet.length <= nextIndex) {
          return {done: true};
        }
      }
    };
  };

  function hasDayOfMonth(dayOfMonth) {
    return _timesheet[dayOfMonth] ? true : false;
  };

  function create(date) {
    var dayOfMonth = date.getDate();
    _timesheet[dayOfMonth] = new TimesheetDayEntry(date);
  };

  function get(dayOfMonth) {
    return _timesheet[dayOfMonth];
  }

}

function TimesheetDayEntry(date, clients) {

  var _date = date;
  var _clients = clients ? clients : {};

  this.add = function(client, duration) {
    if (hasClient(client)) {
      addDurationTo(client, duration);
    } else {
      setDurationFor(client, duration);
    }
  };

  this.clients = function() {
    return _clients;
  };

  this.date = function() {
    return _date;
  };

  this.iterator = function() {
    var nextIndex = 0;
    
    return {
      next: function() {
        if (nextIndex < Object.keys(_clients).length) {
          var client = Object.keys(_clients)[nextIndex++];
          var duration = _clients[client];
          return {value: {clientName: client, duration: duration}, done: false}
        } else {
          return {done: true};
        }
      }
    };
  };

  this.equals = function(other) {
    if (!other) return false;
  };

  function addDurationTo(client, duration) {
    _clients[client] = _clients[client] + duration;
  }

  function setDurationFor(client, duration) {
    _clients[client] = duration;
  }

  function hasClient(client) {
    return _clients[client];
  }
}
