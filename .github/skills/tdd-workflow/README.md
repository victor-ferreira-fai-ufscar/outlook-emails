# TDD Workflow Skill

This skill teaches and guides you through Test-Driven Development (TDD) for implementing new features in the FastAPI Outlook integration project.

## Files in this skill

- **SKILL.md** - Complete TDD guide with Red-Green-Refactor cycle, best practices, and examples
- **test-template.py** - Template for writing TDD tests with fixtures and test patterns
- **implementation-template.py** - Template for implementing code after tests are written
- **README.md** - This file

## Quick Start

### 1. Use the `/tdd-workflow` slash command in Copilot Chat

Type `/tdd-workflow` followed by your feature description:

```
/tdd-workflow I want to add a new /data/export endpoint that exports user data as JSON
```

Copilot will guide you through the TDD workflow.

### 2. Follow the Red-Green-Refactor Cycle

**RED:** Write failing tests
```bash
cp .github/skills/tdd-workflow/test-template.py tests/test_your_feature.py
# Edit test-template.py with your feature tests
pytest tests/test_your_feature.py -v  # Should fail
```

**GREEN:** Implement minimal code
```bash
cp .github/skills/tdd-workflow/implementation-template.py app/routes/your_feature.py
# Edit implementation-template.py to make tests pass
pytest tests/test_your_feature.py -v  # Should pass
```

**REFACTOR:** Clean up code
```bash
# Improve code quality, names, comments
pytest tests/ -v  # Ensure nothing broke
```

### 3. Run all tests before committing

```bash
# Run full test suite
pytest tests/ -v

# With coverage
pytest --cov=app tests/

# Specific test file
pytest tests/test_your_feature.py -v
```

## Key Concepts

### The Red-Green-Refactor Cycle

1. **RED** 🔴: Write tests that describe the feature → tests fail
2. **GREEN** 🟢: Write minimal code to make tests pass → tests pass
3. **REFACTOR** 🔵: Improve code quality → tests still pass

This ensures:
- Code works as expected (tests prove it)
- No over-engineering (write only what's needed)
- Clean, maintainable code (refactored at the end)

### Test Structure

```python
# Setup (arrange)
session_id = authenticated_session

# Action (act)
response = client.get(
    "/endpoint",
    cookies={"local_session_id": session_id}
)

# Verification (assert)
assert response.status_code == 200
```

### Common Test Patterns

**Testing authenticated endpoints:**
```python
def test_endpoint_requires_auth(self):
    response = client.get("/protected-endpoint")
    assert response.status_code == 401

def test_endpoint_with_valid_auth(self, authenticated_session):
    response = client.get(
        "/protected-endpoint",
        cookies={"local_session_id": authenticated_session}
    )
    assert response.status_code == 200
```

**Mocking external APIs:**
```python
from unittest.mock import patch, MagicMock

def test_with_mocked_api(self):
    mock_resp = MagicMock()
    mock_resp.status_code = 200
    mock_resp.json.return_value = {"data": "value"}
    
    with patch("app.utils.requests.get", return_value=mock_resp):
        response = client.get("/endpoint")
    
    assert response.status_code == 200
```

## Best Practices

✅ **DO:**
- Write descriptive test names: `test_endpoint_returns_401_without_auth`
- Test one behavior per test function
- Mock external dependencies (APIs, databases)
- Test both happy path and error cases
- Run full test suite before committing

❌ **DON'T:**
- Skip tests for "quick fixes"
- Mix multiple assertions without clear intent
- Test implementation details instead of behavior
- Write tests after code (that's not TDD!)
- Commit code without running `pytest`

## Project-Specific Notes

### For this FastAPI Outlook Integration:

**Authentication:**
- Use `authenticated_session` fixture for tests requiring auth
- Use `mock.patch("app.utils.requests.get")` for Graph API calls
- Session files are auto-created in `sessions/` directory

**File Structure:**
```
app/
├── main.py          # Router registration only
├── config.py        # Configuration constants
├── utils.py         # Helper functions, Graph API calls
└── routes/
    ├── auth.py      # OAuth endpoints
    ├── profile.py   # Profile endpoints
    └── messages.py  # Email endpoints
```

**Running Tests:**
```bash
# Full test suite
pytest tests/ -v

# Specific test file
pytest tests/test_auth.py -v

# Stop at first failure
pytest -x

# With coverage report
pytest --cov=app tests/
```

## Resources

- [Complete TDD Guide](./SKILL.md)
- [Test Template](./test-template.py)
- [Implementation Template](./implementation-template.py)
- [pytest Documentation](https://docs.pytest.org/)
- [FastAPI Testing](https://fastapi.tiangolo.com/advanced/testing-dependencies/)
- [unittest.mock](https://docs.python.org/3/library/unittest.mock.html)

## Need Help?

1. Read the complete [SKILL.md](./SKILL.md) guide
2. Copy and adapt the templates
3. Follow the Red-Green-Refactor cycle
4. Run tests frequently
5. Ask Copilot: `/tdd-workflow [your feature description]`
