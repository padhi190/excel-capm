# Excel-capm
<img src="images/Calculation.png">
<img src="images/result.png" width = 500>

This Excel model was created to help students understanding of the concept of *Efficient Frontier* - where we can get a set of optimal portfolios that offer the highest expected return for a defined level of risk.
<br>
## Data Preparation
Inputs:
- `Start Date`, `End Date`
- `Period` *(Daily, Weekly, Monthly)*
- `Allow Short Selling`
- `Stock Symbols`
- `Risk free rate return`
- `Degree of Risk Aversion`

<img src="images/data.png">

Click the `Download` button to get the stock data from Yahoo finance.  
<br>
## Covariance calculation

After the data is downloaded, the **Covariance Matrix** of those stocks is automatically calculated using the excel formula.<br>
<img src="images/covariance.png">
This **Covariance Matrix** calculation is a required input for the VBA to calculate the **Efficient Frontier** of theses stocks.
<br>

## Portfolio Optimization

We first assign random weights to the stocks and calculate the portfolio return and standard deviation. When we click the `Calculate Optimal Portfolio` button, the VBA will use Solver to change these weights with the goal of achieving our target return while minimizing our standard deviation and keeping the sum of these weights to 1. 
<br>
<img src="images/solver.png" width = 700>
<br>
The VBA code will calculate the **global minimum variance** of the portfolio as the basis of its calculation. This is the bolded row in the weights data table.<br>
After that, the VBA will calculate the efficient frontier by incrementally increasing the target return of the **global minimum variance** and use the Solver once again to complete our chart data.
<img src="images/Calculation.png">

<br>

## Result
In the result sheet, we can see the graphical interpretation of our portfolio. To get this, we have to input our degree of risk aversion (higher means more risk averse). <br>
In this example, we set the risk aversion to 4. Here we see that our optimal portfolio comprised of 13.6% risk free asset, and 86.4% risky assets. The risky assets are further broken down into each individual stock allocation.
<br>
<img src="images/result.png" width = 600>