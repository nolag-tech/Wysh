

Console.WriteLine("Got here.");

!function(global) {
	global.doOutput = function() {
		Console.WriteLine("Have a nice day.");
	};
}(this);

doOutput();
