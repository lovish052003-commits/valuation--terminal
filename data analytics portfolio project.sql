select
year(order_date) as order_year,
month(order_date) as order_month,
sum(sales_amount) as total_sales,
count(distinct customer_key) as total_customers,
sum(quantity) as total_quantity
from fact_sales
where order_date is not null
group by year(order_date), month(order_date)
order by year(order_date), month(order_date)



select
datetrunc(month,order_date) as order_date,
sum(sales_amount) as total_sales,
count(distinct customer_key) as total_customers,
sum(quantity) as total_quantity
from fact_sales
where order_date is not null
group by datetrunc(month,order_date)
order by datetrunc(month,order_date)


select
format(order_date, 'yyyy-MMM') as order_date,
sum(sales_amount) as total_sales,
count(distinct customer_key) as total_customers,
sum(quantity) as total_quantity
from fact_sales
where order_date is not null
group by format(order_date, 'yyyy-MMM')
order by format(order_date, 'yyyy-MMM')



select
order_date,
total_sales,
sum(total_sales)over (order by order_date) as running_total_sales,
avg(avg_price)over (order by order_date) as moving_average_price
from
(
select
datetrunc(month,order_date) as order_date,
sum(sales_amount) as total_sales,
avg(price) as avg_price
from fact_sales
where order_date is not null
group by datetrunc(month,order_date)
) t

with yearly_product_sales as 
(
select
year(f.order_date) as order_year,
p.product_name,
sum(f.sales_amount) as current_sales
from fact_sales f
left join dim_products p
on f.product_key = p.product_key
where order_date is not null
group by year(f.order_date),
p.product_name
)


select
order_year,
product_name,
current_sales,
avg(current_sales) OVER (partition by product_name) avg_sales,
current_sales - avg(current_sales) OVER (partition by product_name) diff_avg_sales,
case
    when current_sales - avg(current_sales) OVER (partition by product_name) > 0 then 'Above Average'
    when current_sales - avg(current_sales) OVER (partition by product_name) < 0 then 'Below Average'
    else 'Avg'
end avg_change,
lag(current_sales) over (partition by product_name order by order_year) py_sales,
current_sales - lag(current_sales) over (partition by product_name order by order_year) as diff_py_sales,
case
    when current_sales - lag(current_sales) over (partition by product_name order by order_year) > 0 then 'Increase'
    when current_sales - lag(current_sales) over (partition by product_name order by order_year) < 0 then 'Decrease'
    else 'No Change)'
end py_change
from yearly_product_sales
order by product_name, order_year



with category_sales as
(
select
Category,
sum(sales_amount) total_sales
from fact_sales f
left join dim_products p
on p.product_key = f.product_key
group by category
)

select
category,
total_sales,
sum(total_sales) over () as overall_sales,
round((cast(total_sales as float)/sum(total_sales) over () )*100 , 2) as percentage_of_total
from category_sales
order by total_sales desc


with product_segment as 
(
select
product_key,
product_name,
cost,
case 
    when cost < 100 then 'Below 100'
    when cost between 100 and 500 then '100-500'
    when cost between 500 and 1000 then '500-1000'
    else 'Above 1000'
end cost_range
from dim_products
)


select
cost_range,
count(product_key) as total_products
from product_segment
group by cost_range


with customer_spending as
(
select
c.customer_key,
sum(f.sales_amount) as total_spending,
min(order_date) as first_order,
max(order_date) as last_order,
datediff(month, min(order_date) , max(order_date)) as lifespan
from fact_sales f
left join dim_customers c
on f.customer_key = c.customer_key
group by c.customer_key
)


select
customer_segment,
count(customer_key) as total_customers
from
(
select
customer_key,
total_spending,
lifespan,
case 
   when lifespan >= 12 and total_spending > 5000 then 'VIP'
   when lifespan>= 12 and total_spending <= 5000 then 'Regular'
   else 'New'
end customer_segment
from customer_spending
)t
group  by customer_segment
order by total_customers desc;


with base_query as 
(
select
f.order_number,
f.product_key,
f.order_date,
f.sales_amount,
f.quantity,
c.customer_key,
c.customer_number,
concat(c.first_name,' ',c.last_name) as customer_name,
datediff(year, c.birthdate, getdate()) age
from fact_sales f
left join dim_customers c 
on c.customer_key = f.customer_key
where order_date is not null
),


customer_aggregation as
(
select
   customer_key,
   customer_number,
   customer_name,
   age,
   count(distinct order_number) as total_orders,
   sum(sales_amount) as total_sales,
   sum(quantity) as total_quantity,
   count(distinct product_key) as total_products,
   max(order_date) as last_order_date,
   datediff(month, min(order_date),max(order_date)) as lifespan
from base_query
group by
customer_key,
   customer_number,
   customer_name,
   age
   )



select
customer_key,
   customer_number,
   customer_name,
   age,
   case
   when age < 20 then'Under 20'
   when age between 20 and 29 then '20-29'
   when age between 30 and 39 then '30-39'
   when age between 40 and 49 then '40-49'
   else '50 and Above'
end as age_group,

   case 
   when lifespan >= 12 and total_sales> 5000 then 'VIP'
   when lifespan>= 12 and total_sales <= 5000 then 'Regular'
   else 'New'
end customer_segment,
   total_orders,
   last_order_date,
   datediff(month, last_order_date, getdate()) as recency,
   total_sales,
   total_quantity,
   total_products,
   last_order_date,
   lifespan,
   case
      when total_sales = 0 then 0
      else total_sales/total_orders
    end as avg_order_value,

    case 
       when lifespan = 0 then total_sales
       else total_sales/lifespan
    end as avg_monthly_spend

from customer_aggregation





