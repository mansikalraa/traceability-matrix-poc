# Risk Management

## H-001 — Invalid DICOM Upload

### Description

An invalid or incomplete DICOM study may be uploaded.

### Risk

The system may process invalid imaging data.

### Mitigation

Validate the DICOM dataset before processing.

---

## H-002 — Incorrect IVD Measurement

### Description

The measurement engine may calculate an incorrect anatomical measurement.

### Risk

Incorrect measurements may affect downstream planning.

### Mitigation

Validate calculated measurements against reference data.

---

## H-003 — Incorrect Surgical Trajectory

### Description

The trajectory engine may generate an incorrect trajectory.

### Risk

An incorrect trajectory could result in unsafe planning.

### Mitigation

Validate trajectory calculations against approved scenarios.

---

## H-004 — Incorrect Navigation Export

### Description

The exported navigation plan may not represent the approved plan.

### Risk

Incorrect information could be sent to the navigation system.

### Mitigation

Validate exported navigation data against the approved plan.
