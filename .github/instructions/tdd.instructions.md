---
description: "Use when implementing new features, adding endpoints, creating new functions, fixing bugs with tests, refactoring with safety. Apply TDD workflow: write failing tests first (RED), implement minimal code (GREEN), improve code quality (REFACTOR)."
---

# Test-Driven Development (TDD) - Auto-Applied

This instruction automatically applies when you're implementing new features, endpoints, or fixing bugs. **Always follow TDD first.**

## 🔴 RED Phase: Write Failing Tests First

**Before writing any implementation code:**

1. Create a test file: `tests/test_[feature_name].py`
2. Write tests that **describe the desired behavior**
3. Run tests with `pytest tests/test_[feature_name].py -v`
4. **Confirm they FAIL** (they should, the feature doesn't exist yet)

### Test Writing Guidelines

- ✅ Write one test per behavior
- ✅ Use descriptive test names: `test_endpoint_returns_200_on_success()`
- ✅ Mock external dependencies (Graph API, files, databases)
- ✅ Test happy paths AND error cases
- ✅ Test edge cases (empty responses, special characters)

### Quick Start: Copy Test Template

```bash
cp .github/skills/tdd-workflow/test-template.py tests/test_[feature].py
```

Then adapt the template to your feature.

## 🟢 GREEN Phase: Implement Minimal Code

**After tests fail:**

1. Create the implementation file
2. Write **minimal code** to make tests pass
3. Run tests with `pytest tests/test_[feature].py -v`
4. **Confirm they PASS** (all tests green)
5. **Don't over-engineer** at this stage

### Implementation Guidelines

- ✅ Implement only what tests require
- ✅ Don't add extra features not covered by tests
- ✅ Make code readable, but focus on passing tests
- ✅ Register routers in `app/main.py` if adding endpoints

### Quick Start: Copy Implementation Template

```bash
cp .github/skills/tdd-workflow/implementation-template.py app/routes/[feature].py
```

Then adapt the template to your feature.

## 🔵 REFACTOR Phase: Improve Code Quality

**After tests pass:**

1. Improve code readability, naming, and structure
2. Add docstrings and comments
3. Extract reusable functions
4. Run full test suite: `pytest tests/ -v`
5. **Confirm all tests STILL PASS**

### Refactoring Guidelines

- ✅ Only refactor when tests are green
- ✅ Make small changes and test frequently
- ✅ Improve variable names, add comments
- ✅ Extract helper functions
- ✅ Keep logic simple and clear

## 📋 TDD Workflow Checklist

### Phase 1: RED ❌
- [ ] Created test file in `tests/test_[feature].py`
- [ ] Wrote descriptive test cases
- [ ] Mocked external dependencies
- [ ] Tests fail when run: `pytest tests/test_[feature].py -v`

### Phase 2: GREEN ✅
- [ ] Created implementation file
- [ ] Wrote minimal code to pass tests
- [ ] Tests pass: `pytest tests/test_[feature].py -v`
- [ ] Registered router in `app/main.py` (if endpoint)

### Phase 3: REFACTOR 🎨
- [ ] Improved code quality
- [ ] Added docstrings and comments
- [ ] Full test suite passes: `pytest tests/ -v`
- [ ] No test regressions

## 🎯 Common Scenarios

### Adding a New API Endpoint

```python
# TESTS FIRST (tests/test_emails_send.py)
def test_send_email_returns_200_on_success():
    """Test happy path"""
    response = client.post("/emails/send", json={...})
    assert response.status_code == 200

def test_send_email_requires_authentication():
    """Test auth enforcement"""
    response = client.post("/emails/send", json={...})
    assert response.status_code == 401

# THEN IMPLEMENT (app/routes/emails.py)
@router.post("/send")
def send_email(email: EmailRequest, session_id: str):
    # Minimal implementation to pass tests
    ...
```

### Creating a Utility Function

```python
# TESTS FIRST (tests/test_validators.py)
def test_validate_email_accepts_valid():
    assert validate_email("user@example.com") == True

def test_validate_email_rejects_invalid():
    assert validate_email("invalid") == False

# THEN IMPLEMENT (app/utils.py)
def validate_email(email: str) -> bool:
    # Minimal implementation to pass tests
    ...
```

### Fixing a Bug

```python
# TESTS FIRST: Write test that reproduces the bug
def test_profile_endpoint_handles_missing_photo():
    """Regression test for bug where missing photo causes crash"""
    response = client.get("/profile", cookies={"local_session_id": session_id})
    assert response.status_code == 200  # Should not crash

# THEN FIX: Write code to handle the edge case
def get_profile(...):
    profile = fetch_outlook_profile(token)
    profile["photo"] = profile.get("photo", None)  # Handle missing photo
    return profile
```

## 🛠️ Useful Commands

```bash
# Run specific test file
pytest tests/test_your_feature.py -v

# Run with coverage report
pytest --cov=app tests/

# Stop on first failure
pytest -x

# Watch mode (requires pytest-watch)
ptw tests/

# Run full test suite before committing
pytest tests/ -v
```

## 📚 Documentation

For detailed guidance, see `.github/skills/tdd-workflow/`:

- **Quick Reference**: [CHEATSHEET.md](./../skills/tdd-workflow/CHEATSHEET.md)
- **Real Example**: [EXAMPLE.md](./../skills/tdd-workflow/EXAMPLE.md)
- **Complete Guide**: [SKILL.md](./../skills/tdd-workflow/SKILL.md)
- **Project Integration**: [INTEGRATION.md](./../skills/tdd-workflow/INTEGRATION.md)
- **Templates**: [test-template.py](./../skills/tdd-workflow/test-template.py), [implementation-template.py](./../skills/tdd-workflow/implementation-template.py)

## ⚡ Pro Tips

1. **Write tests first** — Even if you think you know what the code should do
2. **One assertion per test** (when possible) — Makes failures clear
3. **Mock external APIs** — Don't call real Microsoft Graph API in tests
4. **Test error cases** — Happy path is only part of the story
5. **Commit frequently** — Each phase (RED, GREEN, REFACTOR) can be a commit

## ❓ FAQ

**Q: Do I HAVE to write tests before code?**
A: For new features and bug fixes, yes. TDD prevents bugs and makes refactoring safe.

**Q: What if tests are hard to write?**
A: That's a sign your code has design problems. Make it more testable.

**Q: Can I skip testing for "quick fixes"?**
A: No. "Quick fixes" often break things. Write a test first, always.

**Q: How many tests do I need?**
A: At minimum: 1 happy path + 1 error case. More coverage = more confidence.

## 🔗 Related Resources

- [GitHub's TDD Guide](https://github.blog/ai-and-ml/github-copilot/github-for-beginners-test-driven-development-tdd-with-github-copilot/)
- [pytest Documentation](https://docs.pytest.org/)
- [FastAPI Testing Guide](https://fastapi.tiangolo.com/advanced/testing-dependencies/)
- [unittest.mock](https://docs.python.org/3/library/unittest.mock.html)

---

**Remember:** Tests written first = code you can trust! 🚀

For interactive help, you can also type `/tdd-workflow [feature description]` in Copilot Chat.
