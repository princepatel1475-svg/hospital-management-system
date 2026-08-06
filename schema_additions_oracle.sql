-- Hospital Management System - Additional Modules (Oracle)
-- Run AFTER schema_oracle.sql: adds Doctors, Billing, and Pharmacy/Medical Store tables
-- Usage: sqlplus system/Prince1475@XE @schema_additions_oracle.sql

-- ============================================
-- SEQUENCES
-- ============================================
CREATE SEQUENCE doctors_seq START WITH 1 INCREMENT BY 1;
CREATE SEQUENCE bills_seq START WITH 1 INCREMENT BY 1;
CREATE SEQUENCE medicines_seq START WITH 1 INCREMENT BY 1;
CREATE SEQUENCE medicine_sales_seq START WITH 1 INCREMENT BY 1;

-- ============================================
-- DOCTORS TABLE
-- (separate from generic `staff` -- carries clinical details:
--  specialization, qualification, consultation fee)
-- ============================================
CREATE TABLE doctors (
    doctor_id          NUMBER PRIMARY KEY,
    first_name         VARCHAR2(50) NOT NULL,
    last_name          VARCHAR2(50) NOT NULL,
    specialization     VARCHAR2(100) NOT NULL,
    qualification      VARCHAR2(100),
    phone              VARCHAR2(15),
    email              VARCHAR2(100) UNIQUE,
    consultation_fee   NUMBER(10,2) NOT NULL,
    availability_days  VARCHAR2(100),   -- e.g. 'Mon,Wed,Fri'
    created_at         TIMESTAMP DEFAULT SYSTIMESTAMP
);

CREATE OR REPLACE TRIGGER trg_doctors_id
BEFORE INSERT ON doctors
FOR EACH ROW
WHEN (NEW.doctor_id IS NULL)
BEGIN
    :NEW.doctor_id := doctors_seq.NEXTVAL;
END;
/

-- ============================================
-- BILLS TABLE (patient billing / invoices)
-- ============================================
CREATE TABLE bills (
    bill_id             NUMBER PRIMARY KEY,
    patient_id          NUMBER NOT NULL,
    appointment_id      NUMBER,
    bill_date           DATE DEFAULT SYSDATE NOT NULL,
    consultation_charge NUMBER(10,2) DEFAULT 0,
    medicine_charge     NUMBER(10,2) DEFAULT 0,
    room_charge         NUMBER(10,2) DEFAULT 0,
    other_charge        NUMBER(10,2) DEFAULT 0,
    total_amount        NUMBER(10,2) GENERATED ALWAYS AS
                         (consultation_charge + medicine_charge + room_charge + other_charge) VIRTUAL,
    payment_status       VARCHAR2(10) DEFAULT 'Pending'
                         CHECK (payment_status IN ('Paid','Pending','Partial')),
    payment_method       VARCHAR2(20) CHECK (payment_method IN ('Cash','Card','UPI','Insurance',NULL)),
    created_at            TIMESTAMP DEFAULT SYSTIMESTAMP,
    CONSTRAINT fk_bill_patient FOREIGN KEY (patient_id)
        REFERENCES patients(patient_id) ON DELETE CASCADE,
    CONSTRAINT fk_bill_appointment FOREIGN KEY (appointment_id)
        REFERENCES appointments(appointment_id) ON DELETE SET NULL
);

CREATE OR REPLACE TRIGGER trg_bills_id
BEFORE INSERT ON bills
FOR EACH ROW
WHEN (NEW.bill_id IS NULL)
BEGIN
    :NEW.bill_id := bills_seq.NEXTVAL;
END;
/

-- ============================================
-- MEDICINES TABLE (pharmacy / medical store inventory)
-- ============================================
CREATE TABLE medicines (
    medicine_id     NUMBER PRIMARY KEY,
    medicine_name   VARCHAR2(100) NOT NULL,
    category        VARCHAR2(50),          -- e.g. Tablet, Syrup, Injection
    unit_price      NUMBER(10,2) NOT NULL,
    stock_quantity  NUMBER DEFAULT 0 NOT NULL,
    expiry_date     DATE,
    created_at      TIMESTAMP DEFAULT SYSTIMESTAMP
);

CREATE OR REPLACE TRIGGER trg_medicines_id
BEFORE INSERT ON medicines
FOR EACH ROW
WHEN (NEW.medicine_id IS NULL)
BEGIN
    :NEW.medicine_id := medicines_seq.NEXTVAL;
END;
/

-- ============================================
-- MEDICINE SALES TABLE (medical store revenue)
-- ============================================
CREATE TABLE medicine_sales (
    sale_id        NUMBER PRIMARY KEY,
    medicine_id    NUMBER NOT NULL,
    patient_id     NUMBER,                 -- nullable: walk-in customers with no patient record
    quantity       NUMBER NOT NULL,
    sale_date      DATE DEFAULT SYSDATE NOT NULL,
    total_amount   NUMBER(10,2) NOT NULL,
    created_at     TIMESTAMP DEFAULT SYSTIMESTAMP,
    CONSTRAINT fk_sale_medicine FOREIGN KEY (medicine_id)
        REFERENCES medicines(medicine_id),
    CONSTRAINT fk_sale_patient FOREIGN KEY (patient_id)
        REFERENCES patients(patient_id) ON DELETE SET NULL
);

CREATE OR REPLACE TRIGGER trg_medicine_sales_id
BEFORE INSERT ON medicine_sales
FOR EACH ROW
WHEN (NEW.sale_id IS NULL)
BEGIN
    :NEW.sale_id := medicine_sales_seq.NEXTVAL;
END;
/

-- Auto-decrement stock on each sale
CREATE OR REPLACE TRIGGER trg_medicine_stock_update
AFTER INSERT ON medicine_sales
FOR EACH ROW
BEGIN
    UPDATE medicines
    SET stock_quantity = stock_quantity - :NEW.quantity
    WHERE medicine_id = :NEW.medicine_id;
END;
/

-- ============================================
-- INDEXES
-- ============================================
CREATE INDEX idx_bills_patient ON bills(patient_id);
CREATE INDEX idx_bills_date ON bills(bill_date);
CREATE INDEX idx_sales_medicine ON medicine_sales(medicine_id);
CREATE INDEX idx_sales_date ON medicine_sales(sale_date);

-- ============================================
-- SAMPLE DATA
-- ============================================
INSERT INTO doctors (first_name, last_name, specialization, qualification, phone, email, consultation_fee, availability_days) VALUES
('Ravi', 'Shah', 'Cardiology', 'MD, DM Cardiology', '9876543210', 'dr.ravi.shah@hospital.com', 500, 'Mon,Wed,Fri');
INSERT INTO doctors (first_name, last_name, specialization, qualification, phone, email, consultation_fee, availability_days) VALUES
('Priya', 'Mehta', 'Pediatrics', 'MD Pediatrics', '9876543211', 'dr.priya.mehta@hospital.com', 400, 'Tue,Thu,Sat');
INSERT INTO doctors (first_name, last_name, specialization, qualification, phone, email, consultation_fee, availability_days) VALUES
('Sanjay', 'Rao', 'Orthopedics', 'MS Ortho', '9876543214', 'dr.sanjay.rao@hospital.com', 600, 'Mon,Tue,Thu,Sat');

INSERT INTO medicines (medicine_name, category, unit_price, stock_quantity, expiry_date) VALUES
('Paracetamol 500mg', 'Tablet', 2.50, 500, DATE '2027-06-30');
INSERT INTO medicines (medicine_name, category, unit_price, stock_quantity, expiry_date) VALUES
('Amoxicillin 250mg', 'Capsule', 5.00, 300, DATE '2026-12-31');
INSERT INTO medicines (medicine_name, category, unit_price, stock_quantity, expiry_date) VALUES
('Cough Syrup', 'Syrup', 45.00, 120, DATE '2027-03-15');
INSERT INTO medicines (medicine_name, category, unit_price, stock_quantity, expiry_date) VALUES
('ORS Sachet', 'Powder', 10.00, 400, DATE '2028-01-01');

INSERT INTO bills (patient_id, appointment_id, consultation_charge, medicine_charge, room_charge, other_charge, payment_status, payment_method) VALUES
(1, 1, 500, 100, 0, 0, 'Paid', 'Card');
INSERT INTO bills (patient_id, appointment_id, consultation_charge, medicine_charge, room_charge, other_charge, payment_status, payment_method) VALUES
(2, 2, 400, 50, 0, 0, 'Pending', NULL);
INSERT INTO bills (patient_id, appointment_id, consultation_charge, medicine_charge, room_charge, other_charge, payment_status, payment_method) VALUES
(3, 3, 500, 0, 1200, 0, 'Paid', 'UPI');

INSERT INTO medicine_sales (medicine_id, patient_id, quantity, sale_date, total_amount) VALUES
(1, 1, 10, TRUNC(SYSDATE), 25.00);
INSERT INTO medicine_sales (medicine_id, patient_id, quantity, sale_date, total_amount) VALUES
(3, 2, 1, TRUNC(SYSDATE), 45.00);
INSERT INTO medicine_sales (medicine_id, patient_id, quantity, sale_date, total_amount) VALUES
(4, NULL, 5, TRUNC(SYSDATE), 50.00);

COMMIT;

-- Sanity checks
SELECT 'Doctors: ' || COUNT(*) FROM doctors;
SELECT 'Bills: ' || COUNT(*) FROM bills;
SELECT 'Medicines: ' || COUNT(*) FROM medicines;
SELECT 'Medicine Sales: ' || COUNT(*) FROM medicine_sales;
SELECT 'Total Billing Revenue: ' || SUM(total_amount) FROM bills;
SELECT 'Total Pharmacy Revenue: ' || SUM(total_amount) FROM medicine_sales;
