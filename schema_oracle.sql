-- Hospital Management System - Oracle Database Schema
-- Connect as: sqlplus system/YourPassword@XE
-- Then run: @schema_oracle.sql

-- ============================================
-- SEQUENCES (Oracle has no AUTO_INCREMENT)
-- ============================================
CREATE SEQUENCE staff_seq START WITH 1 INCREMENT BY 1;
CREATE SEQUENCE patients_seq START WITH 1 INCREMENT BY 1;
CREATE SEQUENCE appointments_seq START WITH 1 INCREMENT BY 1;

-- ============================================
-- STAFF TABLE
-- ============================================
CREATE TABLE staff (
    staff_id      NUMBER PRIMARY KEY,
    first_name    VARCHAR2(50) NOT NULL,
    last_name     VARCHAR2(50) NOT NULL,
    role          VARCHAR2(20) CHECK (role IN ('Doctor','Nurse','Receptionist','Admin','Technician')) NOT NULL,
    department    VARCHAR2(50),
    phone         VARCHAR2(15),
    email         VARCHAR2(100) UNIQUE,
    hire_date     DATE NOT NULL,
    created_at    TIMESTAMP DEFAULT SYSTIMESTAMP
);

-- ============================================
-- PATIENTS TABLE
-- ============================================
CREATE TABLE patients (
    patient_id         NUMBER PRIMARY KEY,
    first_name         VARCHAR2(50) NOT NULL,
    last_name          VARCHAR2(50) NOT NULL,
    date_of_birth      DATE NOT NULL,
    gender             VARCHAR2(10) CHECK (gender IN ('Male','Female','Other')) NOT NULL,
    phone              VARCHAR2(15),
    email              VARCHAR2(100),
    address            VARCHAR2(255),
    blood_group        VARCHAR2(5),
    registration_date  DATE DEFAULT SYSDATE,
    created_at         TIMESTAMP DEFAULT SYSTIMESTAMP
);

-- ============================================
-- APPOINTMENTS TABLE
-- ============================================
CREATE TABLE appointments (
    appointment_id     NUMBER PRIMARY KEY,
    patient_id         NUMBER NOT NULL,
    staff_id           NUMBER NOT NULL,
    appointment_date   DATE NOT NULL,
    appointment_time   VARCHAR2(8) NOT NULL,   -- stored as 'HH24:MI:SS'
    reason             VARCHAR2(255),
    status             VARCHAR2(15) DEFAULT 'Scheduled'
                       CHECK (status IN ('Scheduled','Completed','Cancelled','No-Show')),
    notes              CLOB,
    created_at         TIMESTAMP DEFAULT SYSTIMESTAMP,
    CONSTRAINT fk_appt_patient FOREIGN KEY (patient_id)
        REFERENCES patients(patient_id) ON DELETE CASCADE,
    CONSTRAINT fk_appt_staff FOREIGN KEY (staff_id)
        REFERENCES staff(staff_id)
);

-- ============================================
-- TRIGGERS to auto-populate IDs from sequences
-- (Oracle 21c also supports IDENTITY columns, but
--  sequences+triggers are the classic/portable approach
--  and match typical DBMS coursework.)
-- ============================================
CREATE OR REPLACE TRIGGER trg_staff_id
BEFORE INSERT ON staff
FOR EACH ROW
WHEN (NEW.staff_id IS NULL)
BEGIN
    :NEW.staff_id := staff_seq.NEXTVAL;
END;
/

CREATE OR REPLACE TRIGGER trg_patients_id
BEFORE INSERT ON patients
FOR EACH ROW
WHEN (NEW.patient_id IS NULL)
BEGIN
    :NEW.patient_id := patients_seq.NEXTVAL;
END;
/

CREATE OR REPLACE TRIGGER trg_appointments_id
BEFORE INSERT ON appointments
FOR EACH ROW
WHEN (NEW.appointment_id IS NULL)
BEGIN
    :NEW.appointment_id := appointments_seq.NEXTVAL;
END;
/

-- ============================================
-- INDEXES
-- ============================================
CREATE INDEX idx_patient_name ON patients(last_name, first_name);
CREATE INDEX idx_appointment_date ON appointments(appointment_date);
CREATE INDEX idx_appointment_patient ON appointments(patient_id);
CREATE INDEX idx_appointment_staff ON appointments(staff_id);

-- ============================================
-- SAMPLE DATA
-- ============================================
INSERT INTO staff (first_name, last_name, role, department, phone, email, hire_date) VALUES
('Ravi', 'Shah', 'Doctor', 'Cardiology', '9876543210', 'ravi.shah@hospital.com', DATE '2022-03-15');
INSERT INTO staff (first_name, last_name, role, department, phone, email, hire_date) VALUES
('Priya', 'Mehta', 'Doctor', 'Pediatrics', '9876543211', 'priya.mehta@hospital.com', DATE '2021-07-01');
INSERT INTO staff (first_name, last_name, role, department, phone, email, hire_date) VALUES
('Anita', 'Desai', 'Nurse', 'General', '9876543212', 'anita.desai@hospital.com', DATE '2023-01-10');
INSERT INTO staff (first_name, last_name, role, department, phone, email, hire_date) VALUES
('Karan', 'Patel', 'Receptionist', 'Front Desk', '9876543213', 'karan.patel@hospital.com', DATE '2023-05-20');

INSERT INTO patients (first_name, last_name, date_of_birth, gender, phone, email, address, blood_group) VALUES
('Amit', 'Verma', DATE '1990-05-12', 'Male', '9998887771', 'amit.verma@mail.com', 'Ahmedabad, Gujarat', 'B+');
INSERT INTO patients (first_name, last_name, date_of_birth, gender, phone, email, address, blood_group) VALUES
('Sneha', 'Joshi', DATE '1995-11-23', 'Female', '9998887772', 'sneha.joshi@mail.com', 'Ahmedabad, Gujarat', 'O+');
INSERT INTO patients (first_name, last_name, date_of_birth, gender, phone, email, address, blood_group) VALUES
('Rohan', 'Kapoor', DATE '1988-02-08', 'Male', '9998887773', 'rohan.kapoor@mail.com', 'Gandhinagar, Gujarat', 'A-');

INSERT INTO appointments (patient_id, staff_id, appointment_date, appointment_time, reason, status) VALUES
(1, 1, TRUNC(SYSDATE), '10:00:00', 'Routine checkup', 'Scheduled');
INSERT INTO appointments (patient_id, staff_id, appointment_date, appointment_time, reason, status) VALUES
(2, 2, TRUNC(SYSDATE), '11:30:00', 'Fever and cough', 'Scheduled');
INSERT INTO appointments (patient_id, staff_id, appointment_date, appointment_time, reason, status) VALUES
(3, 1, TRUNC(SYSDATE), '14:00:00', 'Follow-up', 'Completed');

COMMIT;

-- Quick sanity check
SELECT 'Staff: ' || COUNT(*) FROM staff;
SELECT 'Patients: ' || COUNT(*) FROM patients;
SELECT 'Appointments: ' || COUNT(*) FROM appointments;
