# VBA-xla-Black-Scholes
xla functions for Black Scholes (call & put) option price and Option Delta
Save the .xla addin files at an accessible URL and add them to the required workbook

Black Scholes Delta addin contains User-defined public functions for:
1. 'BSMprice' - Black Scholes Merton Options Price (for call & put ; European & American options)
2. 'CND' - Cumulative Normal Distribution
3. 'Option Delta' - Option Delta for Call & Put ; European & American options

'Calc Option Range' addin contains User-defined public functions for:
1. 'OptionRange' - Calculating an option price-band (using the Option's Delta and a user-specified range % multiplier) to avoid spurious orders

These above 4 User-defined functions can be invoked into your workbook like any other existing Functions
The parameter and datatype definitions can be seen via VBA code Editor
