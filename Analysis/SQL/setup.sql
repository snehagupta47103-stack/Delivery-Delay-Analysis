CREATE DATABASE delivery_delay_analysis;

USE delivery_delay_analysis;

CREATE TABLE routes(
    route_id INT PRIMARY KEY,
    route TEXT NOT NULL,
    service_type TEXT NOT NULL
);

INSERT INTO routes (route_id,route,service_type) VALUES
('1','Metro Link','Express'),
('2','City Dash','Express'),
('3','Highway Freight','Standard'),
('4','Rural Feeder','Standard');

CREATE TABLE deliveries(
    record_id INT PRIMARY KEY,
    month TEXT NOT NULL,
    route_id INT NOT NULL,
    hub TEXT NOT NULL,
    promised_days INT NOT NULL,
    actual_days INT NOT NULL,
    FOREIGN KEY (route_id) REFERENCES routes(route_id)
);

INSERT INTO deliveries (record_id,month,route_id,hub,promised_days,actual_days) VALUES
('1','Jan','1','Mumbai','2','2'),
('2','Jan','2','Chennai','3','4'),
('3','Jan','3','Delhi','5','8'),
('4','Jan','4','Mumbai','6','10'),
('5','Feb','1','Chennai','2','5'),
('6','Feb','2','Delhi','3','3'),
('7','Feb','3','Delhi','5','10'),
('8','Feb','4','Chennai','6','7'),
('9','Mar','1','Delhi','2','8'),
('10','Mar','2','Mumbai','3','5'),
('11','Mar','3','Chennai','5','5'),
('12','Mar','4','Mumbai','6','15');


