-- S2a: Total delay_days by service type (highest first)
SELECT r.service_type,
       SUM(CASE WHEN d.actual_days - d.promised_days > 0
                THEN d.actual_days - d.promised_days
                ELSE 0 END) AS total_delay_days
FROM deliveries d
JOIN routes r ON d.route_id = r.route_id
GROUP BY r.service_type
ORDER BY total_delay_days DESC;

-- S2b: Routes whose summed delay_days exceeds 8
SELECT r.route,
       SUM(CASE WHEN d.actual_days - d.promised_days > 0
                THEN d.actual_days - d.promised_days
                ELSE 0 END) AS total_delay_days
FROM deliveries d
JOIN routes r ON d.route_id = r.route_id
GROUP BY r.route
HAVING SUM(CASE WHEN d.actual_days - d.promised_days > 0
                THEN d.actual_days - d.promised_days
                ELSE 0 END) > 8
ORDER BY total_delay_days DESC;

-- S2c: Top two hubs by summed delay_days (ties -> alphabetical hub)
SELECT d.hub,
       SUM(CASE WHEN d.actual_days - d.promised_days > 0
                THEN d.actual_days - d.promised_days
                ELSE 0 END) AS total_delay_days
FROM deliveries d
GROUP BY d.hub
ORDER BY total_delay_days DESC, d.hub ASC
LIMIT 2;

-- S3: Data integrity check (LEFT JOIN) - expected unmatched = 0
SELECT COUNT(*) AS unmatched_keys
FROM deliveries d
LEFT JOIN routes r ON d.route_id = r.route_id
WHERE r.route_id IS NULL;