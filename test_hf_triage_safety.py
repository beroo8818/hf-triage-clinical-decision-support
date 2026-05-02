import os
import importlib.util


def find_app_file():
    """
    Find the main HF triage app file in the same folder as this test file.
    This avoids problems caused by names like:
    hf_triage_app_v35_clinical_safety_FIXED_AUDIT(4).py
    """

    current_folder = os.path.dirname(os.path.abspath(__file__))

    for filename in os.listdir(current_folder):
        if filename.startswith("hf_triage_app") and filename.endswith(".py"):
            return os.path.join(current_folder, filename)

    raise FileNotFoundError(
        "Could not find the main app file. "
        "Make sure the Python app file is in the same folder as this test file "
        "and its name starts with hf_triage_app"
    )


def load_app_module():
    app_file = find_app_file()

    spec = importlib.util.spec_from_file_location("hf_triage_app", app_file)
    app = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(app)

    print(f"Loaded app file successfully: {os.path.basename(app_file)}")
    return app


def test_chest_pain_gate(app):
    snapshot = {
        "chest_pain": "Yes",
        "systolic_bp": "120",
        "heart_rate": "80",
        "respiratory_rate": "18",
        "oxygen_saturation": "97",
    }

    gates = app.evaluate_emergency_gates_snapshot(snapshot)

    assert any("Chest pain" in gate for gate in gates), "Chest pain gate failed"


def test_low_oxygen_gate(app):
    snapshot = {
        "oxygen_saturation": "88",
    }

    gates = app.evaluate_emergency_gates_snapshot(snapshot)

    assert any("SpO2" in gate or "hypoxia" in gate.lower() for gate in gates), "Low oxygen gate failed"


def test_hypotension_gate(app):
    snapshot = {
        "systolic_bp": "85",
    }

    gates = app.evaluate_emergency_gates_snapshot(snapshot)

    assert any("Hypotension" in gate for gate in gates), "Hypotension gate failed"


def test_severe_tachycardia_gate(app):
    snapshot = {
        "heart_rate": "135",
    }

    gates = app.evaluate_emergency_gates_snapshot(snapshot)

    assert any("tachycardia" in gate.lower() or "HR" in gate for gate in gates), "Severe tachycardia gate failed"


def test_invalid_spo2_validation(app):
    snapshot = {
        "oxygen_saturation": "120",
    }

    errors = app.validate_vital_ranges_snapshot(snapshot)

    assert len(errors) > 0, "Invalid SpO2 validation failed"


if __name__ == "__main__":
    app = load_app_module()

    test_chest_pain_gate(app)
    test_low_oxygen_gate(app)
    test_hypotension_gate(app)
    test_severe_tachycardia_gate(app)
    test_invalid_spo2_validation(app)

    print("All safety helper tests passed.")