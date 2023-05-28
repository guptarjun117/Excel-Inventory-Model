# Excel-Inventory-Model
Solved a local home-based baker client’s inventory management problem by designing a comprehensive Excel model to forecast raw materials required based on demand patterns, utilizing macros and VBA ~ Year 2022.

----

**Executive Summary**
The objective of our project is to help our client, Avril, owner of the home-bakery @avril.swt, gain a better overview of her business – including profit maximization and operations optimization. Using historical sales data, operational data, costs data, and data from a team-conducted survey on consumers’ preferences, we developed 4 sub-modules: (1) Demand Forecasting, (2) Price Optimization, (3) Ingredient Inventory Optimization and (4) Equipment Optimization - enabling better pricing, inventory, and equipment investment decisions. We utilized (1) solver, (2) trend/seasonality analysis, (3) queue simulations, and (4) VBA to automate and build an interactive dashboard. Our analysis revealed that as a one-man operation, Avril can deliver orders with a quick turnaround and low utilization rate of her oven. Hence, we determined that a manpower optimization sub-module would not be useful for our client in the coming future and removed it from our model. 

**Project Overview (Appendix 1): Business Aspects** 
Avril is a small-sized home-baker specializing in artisan cakes, earning SGD40,000 – 60,000 annually. With plans to grow her business, Avril needs to understand how sales might grow and how her operational requirements might change. To remain competitive, she is also looking to investigate if her current pricing is optimal for maximizing profits. Hence, this project aims to provide Avril with data-driven insights to make better decisions.

**Project Description** 
* Interactive dashboard to review inventory, and gross profits based on historical and forecasted demand.
* Data-input sheet to update sales orders and time taken to prepare orders, to update demand forecast sub-module.
* Demand forecasting using historical sales, customer survey data and new data inputs from future sales of times of orders to conduct a queue simulation.
* Pricing Optimization using linear projection/ trendline based on survey data on willingness to purchase different cake-types at different price points.
* Ingredients Optimization using forecasted demand and solver.
* Equipment Optimization using historical sales data to conduct a queue simulation, Monte Carlo simulation to evaluate the impact of number of ovens on waiting times, and oven utilization rates.

**Motivation** 
Given the fragmented cakes and pastries market, one cannot compete on quality of their cakes alone. Rather, profit optimization in every aspect of the business is key to remain competitive amidst strong competition for customers.

**Stakeholders**
Avril, founder of @avril.swt, focuses on delivering artisanal commissioned cakes, advertised on Instagram and Tik Tok – where we found and reached out to engage her as a client. We conducted 1 zoom interview to understand Avril’s pain-points, and continued to liaise with her over email, providing bi-weekly updates, and had 1 in-person meeting to test our model to gain insight on our model’s usability and usefulness.

**Deliverables** 

_Outcomes:_
* Excel Spreadsheet model with a dashboard and data-input tab for Avril to interact directly with, showcasing insights from 1 demand forecasting, 1 price optimization and 2 operations optimization sub-modules.
* Simple instructional pdf on navigating the interactive tabs.

**Value Statement:** Our user-friendly, self-sustaining model allows Avril to quickly refer to her operations, forecasted profits and optimal prices, while updating sales data easily, as part of her daily operations. Our demand forecast model will allow Avril to prepare future growth strategies – in marketing, menu development or investments.

**Model Architecture (Appendix 2): Modeling Aspects**
Our model consists of a demand forecasting module and an optimization module consisting of price, ingredient inventory and equipment optimization. A user-friendly dashboard takes in key inputs and displays key results of the model.


_**Demand Forecasting**_
![image](https://github.com/guptarjun117/Excel-Inventory-Model/assets/105283893/c083cb64-7615-4799-907e-cdc087b89119)

_Description:_ 
This is a sub-module that aims to give Avril a holistic view of what demand trends could look like going forward by leveraging on past demand data. The sheet takes the expected cake sales based on historically observed interarrival time that has been adjusted by a demand factor (based on the price optimization module). The daily cake sales are run through a randomizer to stimulate the number of each cake type sold. This module is important in helping Avril prepare for upcoming orders by ensuring there is sufficient inventory on-hand, thereby optimizing capital committed to her working capital, as well as saving her time whilst reducing customer waiting time. 

_Key Assumptions and Dependencies:_
The demand forecast has 2 key assumptions: Firstly, demand is constant throughout the year (i.e., there is no seasonality). This was justified via our chat with Avril where she thinks that other factors (like when she reaches out to customers) is much more important. Secondly, historical order volume is representative of the future, which is justified by the fact that order rates have been largely constant, and Avril does not anticipate any future change. To forecast the type of cake sold on a specified day, we used historical proportion of sale for each cake type and ran a randomizer using the relevant values.

_Dynamic Updating:_
To better cater to Avril’s evolving business, we prepared a “Data Input” tab where she can input daily sales with the press of a button. This new data will be considered in the demand forecasting, with cells in the “Sale Data” tab from row 620 onwards capturing these data points.

**Price Optimization**

_Description:_
![image](https://github.com/guptarjun117/Excel-Inventory-Model/assets/105283893/f5a4c706-7d52-4858-8a9c-1592382d4355)

The aim of this sub-model is to find out the **optimal price** Avril should charge for her 5 types of cakes to **maximize profit** made on each cake.

_Features and Interactions_
This sub-model provides the optimized price for each cake based on the fixed and variable costs and survey results. The dashboard allows the user to add items to her inventory which updates the total costs. 
Two cumulative profit line charts are clearly displayed on the dashboard. Using both original demands from sales data and forecast demand from the demand forecast sub-model. To calculate the profit between the original price and optimal price for both original demand and forecasted demand.

_Analysis_
This sub-model performs a trend line analysis by calculating the willingness to purchase various cakes based on different cake prices. The input data for calculating the willingness trend line is from provided sales data. On each cake, we would select the cake price which gives us the largest profit by subtracting both variable cost and fixed cost from revenue. The results are displayed in Appendix 4.

_Assumptions and Dependencies_
In calculating the optimal price for maximizing profit, we assumed that since the model is based on the survey results collected, the current taste preferences and willingness to pay for each cake will be the same in the future. We are also assuming that the survey data is a good representation of the true demand.


**Equipment Forecasting**

_Description_

![image](https://github.com/guptarjun117/Excel-Inventory-Model/assets/105283893/7725673d-9fcc-4f72-bf52-46b5795fd95d)

This sub-module calculates the **optimal number of ovens** Avril should have to **minimize waiting time**. Waiting time is defined by the difference between the time when the order is ready for collection and when the order comes in. 

_Features and Interactions_
We utilize the demand forecast sub-module here to evaluate the optimal number of ovens Avril should have. The key result of this sub-module was communicated to Avril directly and requires no input on her part. We then calculate the utilization rate of her oven, and the result is fed into the dashboard, so Avril knows when to buy a new oven.

_Assumptions and Dependencies_

In calculating the waiting time, we made the following assumptions. First, Avril starts working at 8am; orders that come in before 8am are processed at 8am and orders that come in after 2pm will be processed the next day. Next, cakes that come out of the oven after 4pm can only be collected at 8am the next day. Lastly, an order takes 10min to process as Avril must prepare the ingredients required, and the cake can only be collected 30min after it comes out of the oven to allow it to cool-down and be packed.

**Ingredients Optimization**

_Description_

![image](https://github.com/guptarjun117/Excel-Inventory-Model/assets/105283893/81b7351c-d8c2-4a2d-b485-af8c5be7bf83)

This sub-module aims to find out the minimum amount Avril must spend on topping up her inventory to ensure she has enough inventory to cover the forecasted demand for the next four weeks. In a home baking business, ingredient cost makes up most of the cost of goods sold, hence we thought this would be essential for the model. Additionally, we chose four weeks to keep the average number of each cake demanded in that month somewhat consistent as central limit theorem dictates the distribution of each cake should approach a normal distribution when sample size is large. If one week was used, the sample size would be small, and the number of each cake demanded will change drastically.

_Features and Interactions_

The sub-module allows Avril to input the minimum inventory she would like to keep on hand in terms of the number of each cake, and utilizes our demand forecast sub-module to solve for the minimum inventory expenditure at a point in time. The current assumption, at Avril’s request, is to solve for the minimum inventory required to service the next 4 weeks of forecasted cake demand with a buffer of 2 of each cake. 
Analysis

This sub-module performs a sensitivity analysis which takes in the number of ovens and orders a day to calculate the change in average waiting time. The results are based on the average of a Monti Carlo simulation (Appendix 3).

We used a Monti Carlo simulation as the average waiting time has a probability distribution that has an inherent uncertainty because the interarrival rate is a large number (2237 minutes) and there is a large range of possible values. By using the Monti Carlo simulation, we eliminate as much uncertainty as possible by estimating all possible outcomes and taking the average. We then calculated the utilization rate of the oven by calculating how much time in an average working day the oven is being used. We found that the utilization rate of the oven is very low at 18.8%, hence suggested that the constraint is not the oven and there is no need to get a new oven for now.


_Assumptions and Dependencies_

We assume that Avril replenishes her inventory every four weeks. This assumption has been sense-checked with Avril. The rationale is many of the ingredients such as fresh fruits are perishable but can last up to 4 to 6 weeks in the freezer. Hence, the model is built around 4 weeks of inventory and not 6 weeks to maximize freshness.


_Analysis_

This sub-module is powered by Solver. The parameters minimize total cost of the monthly inventory expenditure, the new amount of inventory required after replenishing must be more than the minimum inventory required to fulfil the forecasted orders for the next 4 weeks plus two (input variable) of each cake as a buffer, and the quantity to buy must be an integer. We chose to minimize the total cost of the monthly inventory expenditure to minimize the chances of overstocking.


**Conclusions**
We shared our model with Avril to gain her input on whether it provided useful, actionable insights, was user-friendly and easy to implement in her daily operations (Appendix 5).

Avril indicated the ingredient inventory dashboard was the most useful in daily operations. Given Avril’s creativity for exploring the baking of different pastries, we believe the inventory module can be updated as Avril adds new products and expands her ingredients lists.

Avril particularly appreciated the easy-to-navigate interactive dashboard and data-input sheet. While the bulk of her work focuses on baking, re-stocking her inventory, the model can be further developed to pull other levers of growth such as marketing to grow cake demands through accessing social media engagement data, beyond the team’s survey (Appendix 6).

Our analysis revealed there was currently insufficient demand for additional manpower. However, if future marketing successfully drives demand to significantly increase equipment utilization rates and customer waiting times - we can add a manpower optimization module to optimize task allocation and staff scheduling.

The team took ownership over individual workstreams – requiring each member to develop project management skills on top of analysis and excel modelling skills. Our biggest takeaway was learning to tackle any problem, by systematically breaking it down to develop an easy-to-understand methodology.
