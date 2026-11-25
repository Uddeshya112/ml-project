import io

from server import app


def test_health_endpoint():
    client = app.test_client()
    response = client.get("/api/health")

    assert response.status_code == 200
    assert response.get_json() == {"status": "ok"}


def test_samples_endpoint_returns_real_data():
    client = app.test_client()
    response = client.get("/api/samples")
    data = response.get_json()

    assert response.status_code == 200
    assert len(data["jds"]) >= 1
    assert len(data["resumes"]) >= 1


def test_analyze_happy_path_and_ranking():
    client = app.test_client()
    response = client.post("/api/analyze", json={
        "jd": "Python AWS Docker Kubernetes Terraform Linux CI/CD",
        "resumes": [
            {"id": "cloud", "name": "cloud_resume.txt", "text": "Python AWS Docker Kubernetes Terraform Linux CI/CD"},
            {"id": "other", "name": "other_resume.txt", "text": "Marketing analytics spreadsheets presentation"},
        ],
    })
    data = response.get_json()

    assert response.status_code == 200
    assert data["candidates"][0]["name"] == "cloud_resume.txt"
    assert data["candidates"][0]["cosinePct"] >= data["candidates"][1]["cosinePct"]
    assert "jdPrep" in data


def test_analyze_rejects_empty_job_description():
    client = app.test_client()
    response = client.post("/api/analyze", json={"jd": "", "resumes": [{"text": "Python"}]})

    assert response.status_code == 400
    assert "empty" in response.get_json()["error"].lower()


def test_parse_plain_text_upload():
    client = app.test_client()
    response = client.post(
        "/api/parse",
        data={"files": (io.BytesIO(b"Python developer with Docker experience"), "resume.txt")},
        content_type="multipart/form-data",
    )
    data = response.get_json()

    assert response.status_code == 200
    assert data["files"][0]["error"] is None
    assert "Python" in data["files"][0]["text"]
