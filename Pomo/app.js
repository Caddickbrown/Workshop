var countdown;
var timerDisplay = document.querySelector('.timer');
var startTime;
var elapsedSeconds = 0; // Add a variable to track the elapsed seconds

function startTimer() {
  var duration = 25 * 60; // Set the duration to 25 minutes

  clearInterval(countdown); // Clear any existing timers

  if(startTime === undefined){ // Only set the start time if it is undefined
    startTime = Date.now(); // Store the current time
  }

  countdown = setInterval(function() {
    elapsedSeconds++; // Increment the elapsed seconds by 1

    var remainingSeconds = duration - elapsedSeconds; // Calculate the remaining seconds

    var minutes = Math.floor(remainingSeconds / 60);
    var seconds = remainingSeconds % 60;

    if (minutes === -1 && seconds === -1) { // Timer has ended
      clearInterval(countdown);
      playSound(); // Add any desired alarm sound
    } else {
      updateTimer(minutes, seconds);
    }
  },1000);

  updateTimer(minutes, seconds); 
}

function pauseTimer() {
	clearInterval(countdown);
}

function stopTimer() {
	clearInterval(countdown);
	updateTimer(25, 0); // Reset the timer to 25 minutes
	elapsedSeconds = 0; // Reset the elapsed seconds back to zero
	startTime = undefined; // Reset the start time back to undefined
}

function updateTimer(minutes, seconds) {
	var formattedMinutes = minutes.toString().padStart(2, '0');
	var formattedSeconds = seconds.toString().padStart(2, '0');
	timerDisplay.textContent = formattedMinutes + ':' + formattedSeconds;
}

function playSound() {
	// Add code here to play an alarm sound when the timer ends
}