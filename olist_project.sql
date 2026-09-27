-- =========================================================
-- OLIST E-COMMERCE ANALYSIS — FULL SQL SCRIPT
-- Rahul Ballidav
-- Cleaned in Excel/Power Query, loaded and analyzed in MySQL
-- =========================================================


-- =========================================================
-- SECTION 1: DATABASE + TABLES
-- =========================================================

CREATE DATABASE IF NOT EXISTS olist_project;
USE olist_project;

CREATE TABLE orders (
    order_id VARCHAR(50) PRIMARY KEY,
    customer_id VARCHAR(50),
    order_status VARCHAR(20),
    order_purchase_timestamp DATETIME,
    order_approved_at DATETIME,
    order_delivered_carrier_date DATETIME,
    order_delivered_customer_date DATETIME,
    order_estimated_delivery_date DATETIME
);

CREATE TABLE customers (
    customer_id VARCHAR(50) PRIMARY KEY,
    customer_unique_id VARCHAR(50),
    customer_zip_code_prefix VARCHAR(10),
    customer_city VARCHAR(100),
    customer_state VARCHAR(5)
);

CREATE TABLE order_items (
    order_id VARCHAR(50),
    order_item_id INT,
    product_id VARCHAR(50),
    seller_id VARCHAR(50),
    shipping_limit_date DATETIME,
    price DECIMAL(10,2),
    freight_value DECIMAL(10,2)
);

CREATE TABLE payments (
    order_id VARCHAR(50),
    payment_sequential INT,
    payment_type VARCHAR(30),
    payment_installments INT,
    payment_value DECIMAL(10,2)
);

CREATE TABLE products (
    product_id VARCHAR(50) PRIMARY KEY,
    product_category_name VARCHAR(100),
    product_name_lenght INT,
    product_description_lenght INT,
    product_photos_qty INT,
    product_weight_g INT,
    product_length_cm INT,
    product_height_cm INT,
    product_width_cm INT
);

CREATE TABLE sellers (
    seller_id VARCHAR(50) PRIMARY KEY,
    seller_zip_code_prefix VARCHAR(10),
    seller_city VARCHAR(100),
    seller_state VARCHAR(5)
);

CREATE TABLE order_reviews (
    review_id VARCHAR(50),
    order_id VARCHAR(50),
    review_score INT,
    review_comment_title VARCHAR(255),
    review_comment_message TEXT,
    review_creation_date DATETIME,
    review_answer_timestamp DATETIME
);

CREATE TABLE category_translation (
    product_category_name VARCHAR(100) PRIMARY KEY,
    product_category_name_english VARCHAR(100)
);


-- =========================================================
-- SECTION 2: LOADING DATA
-- Every CSV is copied into MySQL's trusted upload folder first:
--   C:/ProgramData/MySQL/MySQL Server 8.0/Uploads/
-- CHARACTER SET latin1 is needed because city/category names
-- have Portuguese accents (e.g. "Bárbara").
-- =========================================================

-- --- orders: dates are DD-MM-YYYY HH:MM text, converted with STR_TO_DATE.
-- Some delivery dates are blank (order was cancelled/undelivered), so
-- each one is checked for blank before converting.
LOAD DATA INFILE 'C:/ProgramData/MySQL/MySQL Server 8.0/Uploads/orders.csv'
INTO TABLE orders
FIELDS TERMINATED BY ',' ENCLOSED BY '"' LINES TERMINATED BY '\n' IGNORE 1 ROWS
(order_id, customer_id, order_status, @p1, @p2, @p3, @p4, @p5)
SET
  order_purchase_timestamp     = CASE WHEN @p1='' THEN NULL ELSE STR_TO_DATE(@p1,'%d-%m-%Y %H:%i') END,
  order_approved_at            = CASE WHEN @p2='' THEN NULL ELSE STR_TO_DATE(@p2,'%d-%m-%Y %H:%i') END,
  order_delivered_carrier_date = CASE WHEN @p3='' THEN NULL ELSE STR_TO_DATE(@p3,'%d-%m-%Y %H:%i') END,
  order_delivered_customer_date= CASE WHEN @p4='' THEN NULL ELSE STR_TO_DATE(@p4,'%d-%m-%Y %H:%i') END,
  order_estimated_delivery_date= CASE WHEN @p5='' THEN NULL ELSE STR_TO_DATE(@p5,'%d-%m-%Y %H:%i') END;

-- --- customers: simple load, no date/number conversion needed.
LOAD DATA INFILE 'C:/ProgramData/MySQL/MySQL Server 8.0/Uploads/customers.csv'
CHARACTER SET latin1
INTO TABLE customers
FIELDS TERMINATED BY ',' ENCLOSED BY '"' LINES TERMINATED BY '\n' IGNORE 1 ROWS;

-- --- order_items: shipping_limit_date has seconds (DD-MM-YYYY HH:MM:SS).
LOAD DATA INFILE 'C:/ProgramData/MySQL/MySQL Server 8.0/Uploads/order_items.csv'
INTO TABLE order_items
FIELDS TERMINATED BY ',' ENCLOSED BY '"' LINES TERMINATED BY '\n' IGNORE 1 ROWS
(order_id, order_item_id, product_id, seller_id, @shipdate, price, freight_value)
SET shipping_limit_date = STR_TO_DATE(@shipdate, '%d-%m-%Y %H:%i:%s');

-- --- payments: simple load.
LOAD DATA INFILE 'C:/ProgramData/MySQL/MySQL Server 8.0/Uploads/payments.csv'
INTO TABLE payments
FIELDS TERMINATED BY ',' ENCLOSED BY '"' LINES TERMINATED BY '\n' IGNORE 1 ROWS;

-- --- products: some numeric columns have blanks, staged first as text,
-- then cast safely (REGEXP checks it's all digits before converting).
CREATE TABLE temp_products_check (
    product_id VARCHAR(50), product_category_name VARCHAR(100),
    raw_name_length VARCHAR(20), raw_desc_length VARCHAR(20), raw_photos_qty VARCHAR(20),
    raw_weight VARCHAR(20), raw_length VARCHAR(20), raw_height VARCHAR(20), raw_width VARCHAR(20)
);
LOAD DATA INFILE 'C:/ProgramData/MySQL/MySQL Server 8.0/Uploads/products.csv'
CHARACTER SET latin1
INTO TABLE temp_products_check
FIELDS TERMINATED BY ',' ENCLOSED BY '"' LINES TERMINATED BY '\n' IGNORE 1 ROWS;

INSERT INTO products
SELECT
  product_id,
  NULLIF(product_category_name, ''),
  CASE WHEN raw_name_length REGEXP '^[0-9]+$' THEN CAST(raw_name_length AS UNSIGNED) END,
  CASE WHEN raw_desc_length  REGEXP '^[0-9]+$' THEN CAST(raw_desc_length  AS UNSIGNED) END,
  CASE WHEN raw_photos_qty   REGEXP '^[0-9]+$' THEN CAST(raw_photos_qty   AS UNSIGNED) END,
  CASE WHEN raw_weight       REGEXP '^[0-9]+$' THEN CAST(raw_weight       AS UNSIGNED) END,
  CASE WHEN raw_length       REGEXP '^[0-9]+$' THEN CAST(raw_length       AS UNSIGNED) END,
  CASE WHEN raw_height       REGEXP '^[0-9]+$' THEN CAST(raw_height       AS UNSIGNED) END,
  CASE WHEN raw_width        REGEXP '^[0-9]+$' THEN CAST(raw_width        AS UNSIGNED) END
FROM temp_products_check;
DROP TABLE temp_products_check;

-- --- sellers: needs latin1 for accented city names.
LOAD DATA INFILE 'C:/ProgramData/MySQL/MySQL Server 8.0/Uploads/sellers.csv'
CHARACTER SET latin1
INTO TABLE sellers
FIELDS TERMINATED BY ',' ENCLOSED BY '"' LINES TERMINATED BY '\n' IGNORE 1 ROWS;

-- --- order_reviews: dates here are in US format M/D/YYYY H:MM:SS AM/PM,
-- different from the other tables, so a different STR_TO_DATE format is used.
-- Staged first since review_score also needed a safe numeric check.
CREATE TABLE temp_reviews_check (
    review_id VARCHAR(50), order_id VARCHAR(50), review_score VARCHAR(10),
    review_comment_title VARCHAR(255), review_comment_message TEXT,
    raw_creation_date VARCHAR(50), raw_answer_timestamp VARCHAR(50)
);
LOAD DATA INFILE 'C:/ProgramData/MySQL/MySQL Server 8.0/Uploads/order_reviews.csv'
CHARACTER SET latin1
INTO TABLE temp_reviews_check
FIELDS TERMINATED BY ',' ENCLOSED BY '"' LINES TERMINATED BY '\n' IGNORE 1 ROWS;

INSERT INTO order_reviews
SELECT
  review_id, order_id,
  CASE WHEN review_score REGEXP '^[0-9]+$' THEN CAST(review_score AS UNSIGNED) END,
  NULLIF(review_comment_title, ''),
  NULLIF(review_comment_message, ''),
  CASE WHEN raw_creation_date='' THEN NULL ELSE STR_TO_DATE(raw_creation_date, '%c/%e/%Y %h:%i:%s %p') END,
  CASE WHEN raw_answer_timestamp='' THEN NULL ELSE STR_TO_DATE(raw_answer_timestamp, '%c/%e/%Y %h:%i:%s %p') END
FROM temp_reviews_check;
DROP TABLE temp_reviews_check;

-- --- category_translation: small lookup table.
LOAD DATA INFILE 'C:/ProgramData/MySQL/MySQL Server 8.0/Uploads/category_translation.csv'
CHARACTER SET latin1
INTO TABLE category_translation
FIELDS TERMINATED BY ',' ENCLOSED BY '"' LINES TERMINATED BY '\n' IGNORE 1 ROWS;


-- --- Check every table loaded correctly
SELECT 'orders' AS tbl, COUNT(*) AS rows_loaded FROM orders
UNION ALL SELECT 'customers', COUNT(*) FROM customers
UNION ALL SELECT 'order_items', COUNT(*) FROM order_items
UNION ALL SELECT 'payments', COUNT(*) FROM payments
UNION ALL SELECT 'products', COUNT(*) FROM products
UNION ALL SELECT 'sellers', COUNT(*) FROM sellers
UNION ALL SELECT 'order_reviews', COUNT(*) FROM order_reviews
UNION ALL SELECT 'category_translation', COUNT(*) FROM category_translation;


-- =========================================================
-- SECTION 3: BUSINESS QUESTIONS
-- =========================================================

-- Q1: Do late deliveries hurt customer reviews?
SELECT
  CASE
    WHEN o.order_delivered_customer_date IS NULL THEN 'Not Delivered'
    WHEN o.order_delivered_customer_date <= o.order_estimated_delivery_date THEN 'On Time'
    ELSE 'Late'
  END AS delivery_status,
  COUNT(DISTINCT o.order_id) AS num_orders,
  ROUND(AVG(r.review_score), 2) AS avg_review_score
FROM orders o
JOIN order_reviews r ON o.order_id = r.order_id
GROUP BY delivery_status
ORDER BY avg_review_score;

-- Q1b: Which state has the most late deliveries?
SELECT
  c.customer_state,
  COUNT(*) AS late_orders,
  ROUND(AVG(r.review_score), 2) AS avg_review_score
FROM orders o
JOIN customers c ON o.customer_id = c.customer_id
JOIN order_reviews r ON o.order_id = r.order_id
WHERE o.order_delivered_customer_date > o.order_estimated_delivery_date
GROUP BY c.customer_state
ORDER BY late_orders DESC
LIMIT 10;

-- Q1c: Which state's customers react worst to lateness? (min 30 late orders)
SELECT
  c.customer_state,
  COUNT(*) AS late_orders,
  ROUND(AVG(r.review_score), 2) AS avg_review_score
FROM orders o
JOIN customers c ON o.customer_id = c.customer_id
JOIN order_reviews r ON o.order_id = r.order_id
WHERE o.order_delivered_customer_date > o.order_estimated_delivery_date
GROUP BY c.customer_state
HAVING COUNT(*) >= 30
ORDER BY avg_review_score ASC
LIMIT 10;

-- Q2: Which product categories bring in the most revenue?
SELECT
    COALESCE(pt.product_category_name_english, 'Unknown') AS category,
    COUNT(DISTINCT oi.order_id) AS orders,
    ROUND(SUM(oi.price), 2) AS revenue
FROM order_items oi
JOIN products p ON oi.product_id = p.product_id
LEFT JOIN category_translation pt ON p.product_category_name = pt.product_category_name
GROUP BY category
ORDER BY revenue DESC
LIMIT 10;

-- Q3: What share of customers order more than once?
SELECT
    COUNT(DISTINCT customer_unique_id) AS unique_customers,
    SUM(order_count > 1) AS repeat_customers,
    ROUND(SUM(order_count > 1) / COUNT(DISTINCT customer_unique_id) * 100, 2) AS repeat_rate_pct
FROM (
    SELECT c.customer_unique_id, COUNT(o.order_id) AS order_count
    FROM customers c
    JOIN orders o ON c.customer_id = o.customer_id
    GROUP BY c.customer_unique_id
) repeat_check;

-- Q4: How do customers pay, and is any of it stuck/unresolved?
SELECT
    payment_type,
    COUNT(*) AS num_payments,
    ROUND(SUM(payment_value), 2) AS total_value
FROM payments
GROUP BY payment_type
ORDER BY total_value DESC;

-- Q5: Do the top sellers by revenue also have good ratings?
SELECT
    s.seller_id,
    ROUND(SUM(oi.price), 2) AS revenue,
    ROUND(AVG(r.review_score), 2) AS avg_rating
FROM order_items oi
JOIN sellers s ON oi.seller_id = s.seller_id
JOIN orders o ON oi.order_id = o.order_id
JOIN order_reviews r ON o.order_id = r.order_id
GROUP BY s.seller_id
ORDER BY revenue DESC
LIMIT 10;

-- Q6a: How many days early do on-time orders arrive, on average?
SELECT
    ROUND(AVG(DATEDIFF(order_estimated_delivery_date, order_delivered_customer_date)), 1)
    AS avg_days_estimate_beat_by
FROM orders
WHERE order_delivered_customer_date IS NOT NULL
  AND order_delivered_customer_date <= order_estimated_delivery_date;

-- Q6b: Bucket every delivered order by how accurate the estimate was.
SELECT
    CASE
        WHEN DATEDIFF(order_estimated_delivery_date, order_delivered_customer_date) >= 10 THEN 'Estimate padded 10+ days'
        WHEN DATEDIFF(order_estimated_delivery_date, order_delivered_customer_date) BETWEEN 1 AND 9 THEN 'Arrived a bit early'
        WHEN DATEDIFF(order_estimated_delivery_date, order_delivered_customer_date) = 0 THEN 'Arrived exactly on estimate'
        ELSE 'Arrived late'
    END AS accuracy_bucket,
    COUNT(*) AS num_orders
FROM orders
WHERE order_delivered_customer_date IS NOT NULL
GROUP BY accuracy_bucket
ORDER BY num_orders DESC;

-- Q6c: Has the estimate gotten more accurate over time (by month)?
SELECT
    DATE_FORMAT(order_purchase_timestamp, '%Y-%m') AS order_month,
    ROUND(AVG(DATEDIFF(order_estimated_delivery_date, order_delivered_customer_date)), 1)
    AS avg_days_early_or_late
FROM orders
WHERE order_delivered_customer_date IS NOT NULL
GROUP BY order_month
ORDER BY order_month;


-- =========================================================
-- SECTION 4: BASIC PRACTICE QUERIES (not part of the write-up,
-- kept here for SQL fluency / interview practice)
-- =========================================================

-- Top 5 cities by number of orders
SELECT c.customer_city, COUNT(DISTINCT o.order_id) AS total_orders
FROM orders o JOIN customers c ON o.customer_id = c.customer_id
GROUP BY c.customer_city ORDER BY total_orders DESC LIMIT 5;

-- Top 5 cities by revenue
SELECT c.customer_city, ROUND(SUM(oi.price), 2) AS revenue
FROM orders o
JOIN customers c ON o.customer_id = c.customer_id
JOIN order_items oi ON o.order_id = oi.order_id
GROUP BY c.customer_city ORDER BY revenue DESC LIMIT 5;

-- Top 10 individual products by revenue
SELECT oi.product_id, COUNT(*) AS times_sold, ROUND(SUM(oi.price), 2) AS revenue
FROM order_items oi GROUP BY oi.product_id ORDER BY revenue DESC LIMIT 10;

-- Top 5 sellers by number of orders
SELECT seller_id, COUNT(DISTINCT order_id) AS total_orders
FROM order_items GROUP BY seller_id ORDER BY total_orders DESC LIMIT 5;

-- Average delivery time by state (in days)
SELECT c.customer_state,
       ROUND(AVG(DATEDIFF(o.order_delivered_customer_date, o.order_purchase_timestamp)), 1) AS avg_days
FROM orders o JOIN customers c ON o.customer_id = c.customer_id
WHERE o.order_delivered_customer_date IS NOT NULL
GROUP BY c.customer_state ORDER BY avg_days DESC LIMIT 10;

-- Order status breakdown
SELECT order_status, COUNT(*) AS total FROM orders GROUP BY order_status ORDER BY total DESC;

-- Review score distribution
SELECT review_score, COUNT(*) AS total FROM order_reviews GROUP BY review_score ORDER BY review_score;

-- Average order value overall
SELECT ROUND(SUM(price) / COUNT(DISTINCT order_id), 2) AS avg_order_value FROM order_items;
