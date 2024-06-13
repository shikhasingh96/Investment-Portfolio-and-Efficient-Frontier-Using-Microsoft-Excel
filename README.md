## Optimal Portfolio and Efficient Frontier Using Microsoft Excel

## Project Description
This project focuses on developing an optimal investment portfolio and mapping the Efficient Frontier using Microsoft Excel. The Efficient Frontier is a key concept in Modern Portfolio Theory, representing the set of optimal portfolios that offer the highest expected return for a given level of risk or the lowest risk for a given level of expected return. The project involves collecting historical stock data, calculating returns and risks, and utilizing Excel tools to optimize the portfolio and visualize the Efficient Frontier.

## Summary
Portfolio managers and investors aim to achieve the best trade-off between risk and return when constructing portfolios from a set of assets. The risk/return profile of a portfolio changes with its composition. Portfolios that maximize return for a given risk or minimize risk for a given return are termed optimal portfolios. The set of these optimal portfolios is known as the Efficient Frontier.

In this project, we develop an optimal investment portfolio and map the Efficient Frontier using historical data of US equities downloaded from Yahoo! Finance. The stock prices from 2014 to 2024 for twelve companies across six sectors (Technology, Healthcare, Financial, Consumer Commodities, Real Estate, and Energy) are used to ensure diversification. The project is organized as follows:

1. Review of Concepts: Expected return, risk, diversification, and capital allocation.
2. Data Collection: Gathering historical stock prices.
3. Data Analysis: Calculating monthly returns, statistical measures, and covariance matrices.
4. Optimization Model: Creating an Excel-based model using the Solver add-in to optimize the portfolio.
5. Conclusion: Summarizing the findings and optimal portfolio allocation.


## Main Chapter
- Data Collection
The data for this project was collected from Yahoo! Finance, covering ten years of monthly stock prices (2014-2024) for twelve companies across six sectors:

SECTORS	COMPANIES

Health Care	Abbott Laboratories (ABT)

Pfizer Inc. (PFE)

Technology	Apple Inc. (AAPL)

Microsoft Corporation (MSFT)

Consumer Commodities	McDonald's Corporation (MCD)

Starbucks Corporation (SBUX)

Finance	Visa Inc. (V)

Mastercard Incorporated (MA)

Energy	Shell plc (SHEL)

Exxon Mobil Corporation (XOM)

Real Estate	Realty Income Corporation (O)


Public Storage (PSA)

- Data Analysis
The collected data was processed to calculate monthly returns, average monthly returns, annual returns, and variances. Covariance matrices were computed using excess returns to understand how the stocks' returns move together. The Excel functions MMULT and TRANSPOSE were used extensively for matrix operations.
## Key Concepts

Expected Return: The forecast gain or loss on an investment over a period. Calculated as the weighted average of the expected returns on the individual stocks in the portfolio.

Risk: Measured by the standard deviation of returns. Calculated using the variance-covariance matrix of the stock returns.
![image](https://github.com/shikhasingh96/Investment-Portfolio-and-Efficient-Frontier-Using-Microsoft-Excel/assets/136284820/4cf200b6-226f-49a2-a9a5-3c348b30891e)

![image](https://github.com/shikhasingh96/Investment-Portfolio-and-Efficient-Frontier-Using-Microsoft-Excel/assets/136284820/32d3bdd1-1c09-4079-94e3-783adcaa32bc)

Diversification: Reducing risk by investing in a variety of assets.

Capital Allocation: The optimal mix of weights for the assets in the portfolio to achieve the best risk-return trade-off.

- Optimization Model
An Excel-based optimization model was created with the following decision variables and constraints:

Decision Variables: Investments in each of the twelve companies.

Constraints:
Minimum expected return of 10%.
Maximum risk of 13%.
Total investment distribution summing to 1.
Non-negativity constraint for investments.
The objective function was to maximize the Sharpe ratio, which measures the excess return per unit of risk.

## Optimization Models
Three models were created and analyzed:

Model 1: Only the sum of the decision variable constraint was applied. This model showed a high risk and low diversification.
![image](https://github.com/shikhasingh96/Investment-Portfolio-and-Efficient-Frontier-Using-Microsoft-Excel/assets/136284820/949ca167-cf7d-4168-b892-73005cc91793)

Model 2: Applied all constraints. This model provided a well-diversified portfolio with a Sharpe ratio of 0.917, suggesting optimal risk-adjusted returns.
![image](https://github.com/shikhasingh96/Investment-Portfolio-and-Efficient-Frontier-Using-Microsoft-Excel/assets/136284820/de86cff1-d7ae-435c-9e92-ca6f818667cb)
![image](https://github.com/shikhasingh96/Investment-Portfolio-and-Efficient-Frontier-Using-Microsoft-Excel/assets/136284820/ab09f30d-68b6-4485-abac-5faa8514f26f)


Model 3: Aimed to maximize expected returns with adjusted constraints. This model showed higher expected returns but did not satisfy all constraints, making it less feasible.
![image](https://github.com/shikhasingh96/Investment-Portfolio-and-Efficient-Frontier-Using-Microsoft-Excel/assets/136284820/aaab780a-eb3a-4777-bb83-0747d7ab6a09)

## Sensitivity Analysis

A sensitivity analysis was conducted on Model 2 using one-way and two-way analyses to study the impact of changes in minimum expected return and maximum risk on the Sharpe ratio. The analysis revealed the range of feasible solutions and their sensitivity to variations in constraints.
- One-wayanalyses
![image](https://github.com/shikhasingh96/Investment-Portfolio-and-Efficient-Frontier-Using-Microsoft-Excel/assets/136284820/acd93284-3919-4c70-b9e1-3d8871e12668)
![image](https://github.com/shikhasingh96/Investment-Portfolio-and-Efficient-Frontier-Using-Microsoft-Excel/assets/136284820/4386d044-d536-44cd-bf05-813f59bec69f)
As can be seen from the sensitivity table Figure 4 and associated chart Figure5, for the maximum expected return between 1 and 16 approximately, there is an optimal solution (feasible solutions) is sensitive to variation of minimum expected return. At the same time, when the minimum expected Return is 16 or more, it becomes insensitive, not feasible solution for the return, or otherwise, insensitive to variations of the minimum expected return between 17 and 25.
- Two-way analyses
![image](https://github.com/shikhasingh96/Investment-Portfolio-and-Efficient-Frontier-Using-Microsoft-Excel/assets/136284820/18916779-860b-4d27-99a1-f407e3bee720)
![image](https://github.com/shikhasingh96/Investment-Portfolio-and-Efficient-Frontier-Using-Microsoft-Excel/assets/136284820/f46a012c-bd71-4a4b-ac27-19d71129bb7b)
1.	Maximum Risk: For the two ways sensitivity analyses of parameter Maximum risk and minimum return between 0 to 25 with an increment of one it is obvious that from 0 to 12 the maximum risk is insensitive and there is no feasible solution. But it becomes sensitive at 13 (sharp ratio – 0.917), then it starts increasing till 17, however after that it becomes insensitive to the variation in sharp ratio.	
2.	Minimum Expected Return: For every change in expected return between 0 to 17, there’s no feasible solution. After day 17 to 25, for every price expected return, the solution, the sharp ratio remains same 1.15. 
## Conclusion

The project successfully demonstrated the process of creating an optimal investment portfolio using Excel. Model 2 was identified as the best portfolio, balancing risk and return while satisfying all constraints. The optimized portfolio suggested investments in nine out of twelve companies, ensuring diversification and protecting against downside risks. The Sharpe ratio of 0.917 indicated favorable risk-adjusted performance. Overall, the project highlighted the practical application of Modern Portfolio Theory in achieving financial goals through well-constructed and diversified investment portfolios.

