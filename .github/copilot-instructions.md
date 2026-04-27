---
name: TDD-First Development Culture
description: "Global instruction for all Copilot interactions in this workspace. Always recommend and guide TDD (Test-Driven Development) when implementing features, creating endpoints, or fixing bugs."
---

# 🧪 TDD-First Development Culture

This is a **TDD-first project**. All new features, endpoints, and bug fixes should follow the **Red-Green-Refactor cycle**.

## 🎯 When To Apply TDD

Always use TDD when:
- ✅ **Implementing new endpoints** or routes
- ✅ **Creating new utility functions**
- ✅ **Adding features** to existing code
- ✅ **Fixing bugs** with test-driven approach
- ✅ **Refactoring** existing code

## 🔴 RED: Write Failing Tests

**Step 1:** Before writing any code, write tests that describe the desired behavior.

```python
# tests/test_my_feature.py
def test_my_endpoint_returns_200():
    response = client.get("/my-endpoint")
    assert response.status_code == 200

def test_my_endpoint_requires_auth():
    response = client.get("/my-endpoint")
    assert response.status_code == 401
```

**Step 2:** Run tests and confirm they **FAIL**
```bash
pytest tests/test_my_feature.py -v  # Should show: FAILED
```

## 🟢 GREEN: Implement Minimal Code

**Step 3:** Write minimal code to pass the tests

```python
# app/routes/my_feature.py
@router.get("/my-endpoint")
def my_endpoint(session_id: str = Cookie(None)):
    if not session_id:
        raise HTTPException(status_code=401)
    return {"status": "ok"}
```

**Step 4:** Run tests and confirm they **PASS**
```bash
pytest tests/test_my_feature.py -v  # Should show: PASSED
```

## 🔵 REFACTOR: Improve Code Quality

**Step 5:** Improve code while keeping tests passing

```python
@router.get("/my-endpoint", tags=["features"])
def my_endpoint(session_id: str = Cookie(None)) -> dict:
    """Get my endpoint. Requires authentication."""
    if not session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    return {"status": "ok", "message": "Feature working"}
```

**Step 6:** Run full test suite to ensure no regressions
```bash
pytest tests/ -v  # All tests should pass
```

## 📚 Quick Reference

| Phase | Goal | Command | Expected Result |
|-------|------|---------|-----------------|
| 🔴 RED | Write tests | `pytest tests/test_feature.py -v` | ❌ FAILED |
| 🟢 GREEN | Pass tests | `pytest tests/test_feature.py -v` | ✅ PASSED |
| 🔵 REFACTOR | Improve code | `pytest tests/ -v` | ✅ All PASSED |

## 🎓 Learning Resources

- **Quick Start:** Read [.github/skills/tdd-workflow/CHEATSHEET.md](./.github/skills/tdd-workflow/CHEATSHEET.md)
- **Real Example:** Follow [.github/skills/tdd-workflow/EXAMPLE.md](./.github/skills/tdd-workflow/EXAMPLE.md)
- **Complete Guide:** Study [.github/skills/tdd-workflow/SKILL.md](./.github/skills/tdd-workflow/SKILL.md)
- **Project Help:** Read [.github/skills/tdd-workflow/INTEGRATION.md](./.github/skills/tdd-workflow/INTEGRATION.md)

## 💡 Key Principles

1. **Tests First**: Define behavior before implementation
2. **Minimal Code**: Write only code needed to pass tests
3. **Safety Net**: Tests catch regressions when refactoring
4. **Clear Design**: Tests guide architecture, not vice versa
5. **Living Docs**: Tests document how to use your code

## ⚡ How Copilot Will Help

When you ask to implement a feature, Copilot will:

1. **Suggest test cases** to write first
2. **Provide test templates** matching your project patterns
3. **Guide implementation** that satisfies tests
4. **Recommend refactorings** that keep tests passing
5. **Validate coverage** with full test suite

## 🚀 Example Workflow

**You say:** "Add a new endpoint POST /emails/schedule to schedule emails for later"

**Copilot responds:**
```
Following TDD approach...

PHASE 1: RED - Write Failing Tests
┌─────────────────────────────────┐
- Test: endpoint exists (POST /emails/schedule)
- Test: requires authentication
- Test: validates email data
- Test: returns 201 on success
- Test: handles invalid input (400)

Run: pytest tests/test_emails_schedule.py -v
Result: All tests FAIL ❌
└─────────────────────────────────┘

PHASE 2: GREEN - Implement Minimal Code
...
```

## 🎯 Automation & Auto-Apply

This instruction is **automatically applied** to:
- All feature implementation requests
- All endpoint creation requests
- All bug fix requests
- All refactoring work

You don't need to ask for TDD—Copilot will guide you automatically.

## 💬 When to Use Slash Commands

For **interactive guidance** through the full TDD cycle:
```
/tdd-workflow Add a new endpoint to fetch user's calendar events
```

Copilot will walk through each phase step-by-step.

## ❓ FAQ

**Q: Can I skip TDD for small changes?**
A: TDD takes 10% more time upfront but saves 5x time in debugging. Skip it only for trivial changes (typos, comments).

**Q: What if I don't know what tests to write?**
A: Copilot will help! Say: "I want to add [feature], help me write tests first"

**Q: Do tests slow me down?**
A: Initially yes, but tests compound: each feature adds regression protection. You move faster long-term.

**Q: How do I run tests?**
```bash
pytest tests/ -v                    # All tests
pytest tests/test_feature.py -v     # Specific file
pytest --cov=app tests/             # With coverage
```

## 🔧 Project Testing Commands

```bash
# Install dependencies
uv sync

# Run all tests
pytest tests/ -v

# Run specific test file
pytest tests/test_unit.py -v

# Run with coverage report
pytest --cov=app tests/

# Stop at first failure
pytest -x

# Watch tests on file change (requires pytest-watch)
ptw tests/
```

## 📝 Commit Strategy

After each phase, commit:

```bash
# After RED (tests only)
git add tests/test_feature.py
git commit -m "test: add tests for [feature]"

# After GREEN (implementation)
git add app/routes/feature.py app/main.py
git commit -m "feat: implement [feature]"

# After REFACTOR (improvements)
git add app/routes/feature.py
git commit -m "refactor: improve [feature] code quality"
```

## 🎓 Continuous Improvement

Each feature you build with TDD:
- Makes the codebase more trustworthy
- Creates regression protection
- Documents expected behavior
- Speeds up future refactoring
- Builds team confidence

---

## 🔗 Direct Links

| Resource | Path |
|----------|------|
| TDD Skill Files | `.github/skills/tdd-workflow/` |
| Test Instructions | `.github/instructions/tdd.instructions.md` |
| This File | `.github/copilot-instructions.md` |
| Project README | `README.md` |

---

**Welcome to TDD-first development! 🚀**

Every test written makes your code more trustworthy.
Every test passed brings confidence to refactoring.
Every test is a promise to future maintainers.
