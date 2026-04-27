# TDD Skill Integration Guide

## 📍 Location

This skill is located in `.github/skills/tdd-workflow/` and is automatically discovered by VS Code Copilot.

## 🎯 When to Use This Skill

Use this skill whenever you want to:

- ✅ Add a new API endpoint
- ✅ Implement a new feature or function
- ✅ Extend existing functionality
- ✅ Fix a bug (by writing a test first)
- ✅ Refactor with confidence

**Do NOT skip TDD for:**
- "Quick fixes" (they often break things)
- Bug fixes (write a test to reproduce first)
- Features you're unsure about (TDD clarifies requirements)

## 🚀 Using the Skill in Copilot Chat

### Method 1: Direct Slash Command (Recommended)

In Copilot Chat, type:
```
/tdd-workflow [feature description]
```

Examples:
```
/tdd-workflow Add POST /emails/send endpoint to send emails

/tdd-workflow Create a new utility function to validate email addresses

/tdd-workflow Add caching to the /profile endpoint for performance
```

Copilot will guide you through the entire Red-Green-Refactor cycle.

### Method 2: Manual Step-by-Step

1. Read [CHEATSHEET.md](./CHEATSHEET.md) for quick reference
2. Copy test template: `cp .github/skills/tdd-workflow/test-template.py tests/test_your_feature.py`
3. Write your tests
4. Copy implementation template: `cp .github/skills/tdd-workflow/implementation-template.py app/routes/your_feature.py`
5. Write your implementation
6. Refactor if needed

### Method 3: Detailed Example

Follow the step-by-step example in [EXAMPLE.md](./EXAMPLE.md) for a complete walkthrough.

## 📁 Skill Files Overview

| File                                                       | Purpose                                                               | Read Time |
| ---------------------------------------------------------- | --------------------------------------------------------------------- | --------- |
| [SKILL.md](./SKILL.md)                                     | Complete TDD guide with theory, best practices, and detailed examples | 15 min    |
| [README.md](./README.md)                                   | Overview of the skill, file structure, and quick links                | 5 min     |
| [CHEATSHEET.md](./CHEATSHEET.md)                           | Quick reference card for daily use                                    | 3 min     |
| [EXAMPLE.md](./EXAMPLE.md)                                 | Real-world step-by-step example (creating `/data/summary` endpoint)   | 10 min    |
| [test-template.py](./test-template.py)                     | Template for writing tests - copy and adapt                           | -         |
| [implementation-template.py](./implementation-template.py) | Template for implementing code - copy and adapt                       | -         |

## 🔄 TDD Workflow in This Project

### Project Structure

```
app/
├── main.py          # Only router registration (20 lines)
├── config.py        # Configuration constants
├── utils.py         # Helper functions, Graph API calls
└── routes/          # Individual route modules
    ├── auth.py      # OAuth endpoints
    ├── profile.py   # Profile endpoints
    ├── messages.py  # Email endpoints
    ├── health.py    # Health check
    └── [new_feature.py]  # ← Your new features here

tests/
├── test_unit.py        # Unit tests for utils.py
├── test_integration.py # Integration tests for endpoints
└── test_[new_feature].py  # ← Your new feature tests here
```

### Adding a New Endpoint

#### Step 1: Create Test File

```bash
cp .github/skills/tdd-workflow/test-template.py tests/test_your_endpoint.py
```

Edit `tests/test_your_endpoint.py`:
- Replace `new_endpoint` with your endpoint name
- Replace `/new-endpoint` with your endpoint path
- Add your test cases (happy path + error cases)
- Run: `pytest tests/test_your_endpoint.py -v`
- Expected: ❌ All fail

#### Step 2: Create Implementation File

```bash
cp .github/skills/tdd-workflow/implementation-template.py app/routes/your_endpoint.py
```

Edit `app/routes/your_endpoint.py`:
- Replace endpoint function names with yours
- Implement minimal code to pass tests
- Edit `app/main.py` to register the router:
  ```python
  from app.routes import your_endpoint
  app.include_router(your_endpoint.router)
  ```
- Run: `pytest tests/test_your_endpoint.py -v`
- Expected: ✅ All pass

#### Step 3: Refactor if Needed

```bash
# Improve code quality, variable names, docstrings, etc.
pytest tests/ -v
# Expected: ✅ All tests still pass
```

#### Step 4: Commit with Clear History

```bash
git add tests/test_your_endpoint.py
git commit -m "test: add tests for [feature]"

git add app/routes/your_endpoint.py app/main.py
git commit -m "feat: implement [feature]"

# Optional refactoring commit
git add app/routes/your_endpoint.py
git commit -m "refactor: improve [feature] code quality"
```

### Running Tests

```bash
# Full test suite
pytest tests/ -v

# Specific test file
pytest tests/test_your_endpoint.py -v

# Specific test function
pytest tests/test_your_endpoint.py::test_function_name -v

# With coverage report
pytest --cov=app tests/

# Stop at first failure
pytest -x

# Watch mode (requires pytest-watch)
ptw tests/
```

## 📊 Expected Test Coverage

After adding a new endpoint, you should have tests for:

- ✅ **Happy Path**: Endpoint works correctly with valid input
- ✅ **Authentication**: Endpoint rejects unauthenticated requests (401)
- ✅ **API Errors**: Endpoint handles external API failures (403, 500, etc.)
- ✅ **Data Validation**: Endpoint handles invalid input gracefully
- ✅ **Edge Cases**: Empty responses, special characters, etc.

Example test coverage for an endpoint:
```
test_endpoint_returns_200_on_success          ✅
test_endpoint_returns_required_fields         ✅
test_endpoint_requires_authentication         ✅
test_endpoint_rejects_invalid_session         ✅
test_endpoint_handles_api_error               ✅
test_endpoint_with_empty_response             ✅
test_endpoint_with_special_characters         ✅
```

## 🔧 Customizing Templates

The templates in this skill are starter templates. Adapt them to your needs:

### Modifying Test Template

The [test-template.py](./test-template.py) includes:
- Fixture for authenticated sessions
- Fixture for temporary directories (auto-cleanup)
- Happy path tests (2 examples)
- Error case tests (3 examples)
- Edge case tests (2 examples)

Remove or add test functions as needed for your feature.

### Modifying Implementation Template

The [implementation-template.py](./implementation-template.py) includes:
- FastAPI router setup
- Example endpoint
- Helper function patterns
- Comments showing where to add logic

Keep the structure but replace endpoint logic with yours.

## 📚 Learning Resources

### Within This Skill

1. Start with [CHEATSHEET.md](./CHEATSHEET.md) (3 min read)
2. Follow [EXAMPLE.md](./EXAMPLE.md) step-by-step (10 min)
3. Read [SKILL.md](./SKILL.md) for deep knowledge (15 min)
4. Use templates and this guide as reference

### External Resources

- [GitHub's TDD Guide](https://github.blog/ai-and-ml/github-copilot/github-for-beginners-test-driven-development-tdd-with-github-copilot/)
- [pytest Documentation](https://docs.pytest.org/)
- [FastAPI Testing Guide](https://fastapi.tiangolo.com/advanced/testing-dependencies/)
- [unittest.mock](https://docs.python.org/3/library/unittest.mock.html)

## ✨ Best Practices in This Project

### 1. Keep Tests Focused

```python
# ❌ Too broad
def test_endpoint():
    response = client.get("/endpoint")
    assert response.status_code == 200
    assert "field" in response.json()
    assert response.json()["field"] == "value"

# ✅ Focused
def test_endpoint_returns_200():
    assert response.status_code == 200

def test_endpoint_returns_expected_fields():
    assert "field" in response.json()
```

### 2. Always Mock External APIs

This project calls Microsoft Graph API. Always mock in tests:

```python
from unittest.mock import patch, MagicMock

mock_resp = MagicMock()
mock_resp.status_code = 200
mock_resp.json.return_value = {"data": "mocked"}

with patch("app.utils.requests.get", return_value=mock_resp):
    response = client.get("/endpoint", cookies={"local_session_id": session_id})
```

### 3. Test Authentication First

Most endpoints require `local_session_id` cookie:

```python
def test_endpoint_requires_auth():
    """Always test that auth is enforced."""
    response = client.get("/endpoint")
    assert response.status_code == 401

def test_endpoint_with_valid_session(authenticated_session):
    """Then test with valid auth."""
    response = client.get("/endpoint", cookies={"local_session_id": authenticated_session})
    assert response.status_code == 200
```

### 4. Use Fixtures for Common Setup

```python
@pytest.fixture
def authenticated_session(tmp_path):
    """Reusable fixture for authenticated tests."""
    # Setup code
    return session_id
```

This avoids repeating setup code across many tests.

## 🐛 Troubleshooting

### Issue: Tests pass but endpoint doesn't work

**Cause**: Tests might be inadequate or mocking too much.
**Solution**: 
- Test with real (but mocked) data
- Include integration tests, not just unit tests
- Run full test suite: `pytest tests/ -v`

### Issue: Can't import your new router in main.py

**Cause**: Router not created or wrong module path.
**Solution**:
1. Verify file exists: `ls app/routes/your_file.py`
2. Check import path: `from app.routes import your_file`
3. Check router name: `app.include_router(your_file.router)`

### Issue: Tests pass locally but fail in CI

**Cause**: Usually mocking issues or env variables.
**Solution**:
- Use `tmp_path` fixture for file operations
- Don't depend on .env variables in tests
- Mock all external API calls

### Issue: Test takes too long to run

**Cause**: Making real API calls or file I/O.
**Solution**:
- Mock `requests.get()` calls
- Use `tmp_path` instead of real files
- Keep tests focused and small

## 🎯 Quick Navigation

- **New to TDD?** → Read [CHEATSHEET.md](./CHEATSHEET.md) first
- **Want an example?** → Follow [EXAMPLE.md](./EXAMPLE.md)
- **Need deep knowledge?** → Read [SKILL.md](./SKILL.md)
- **Starting implementation?** → Use templates + [README.md](./README.md)
- **Help with Copilot?** → Type `/tdd-workflow [description]` in chat

---

**Remember:** Tests written first = code you can trust! 🚀
