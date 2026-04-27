# 📚 TDD Skill - Complete Documentation Index

Welcome to the TDD (Test-Driven Development) Skill! This folder contains everything you need to implement features using TDD methodology.

## 🗺️ Navigation Guide

**Choose your starting point based on your needs:**

### 🚀 I want to start RIGHT NOW (5 minutes)
1. Read [CHEATSHEET.md](./CHEATSHEET.md) - Quick reference card
2. Type `/tdd-workflow [your feature description]` in Copilot Chat
3. Follow Copilot's guidance through Red-Green-Refactor cycle

### 📖 I want to learn by example (15 minutes)
1. Read [CHEATSHEET.md](./CHEATSHEET.md) for quick overview
2. Follow [EXAMPLE.md](./EXAMPLE.md) - Complete step-by-step walkthrough
3. Use the example as a template for your own features

### 🎓 I want to understand TDD deeply (30 minutes)
1. Start with [README.md](./README.md) - Overview and key concepts
2. Read [SKILL.md](./SKILL.md) - Complete TDD guide with theory
3. Follow [EXAMPLE.md](./EXAMPLE.md) - Practical application
4. Use [CHEATSHEET.md](./CHEATSHEET.md) as daily reference

### 🔧 I want to integrate this into my workflow (10 minutes)
1. Read [INTEGRATION.md](./INTEGRATION.md) - How to use in this project
2. Copy and adapt the templates
3. Start writing your first TDD feature

---

## 📄 File Directory

### 📋 Documentation Files

| File                                   | Purpose                                                    | Read Time | Best For                       |
| -------------------------------------- | ---------------------------------------------------------- | --------- | ------------------------------ |
| **[CHEATSHEET.md](./CHEATSHEET.md)**   | Quick reference card for daily use                         | 3 min     | Quick lookups, common patterns |
| **[README.md](./README.md)**           | Skill overview, key concepts, best practices               | 5 min     | Understanding TDD fundamentals |
| **[SKILL.md](./SKILL.md)**             | Complete TDD guide with examples and patterns              | 15 min    | Deep understanding of TDD      |
| **[EXAMPLE.md](./EXAMPLE.md)**         | Real-world step-by-step example (`/data/summary` endpoint) | 10 min    | Learning by doing              |
| **[INTEGRATION.md](./INTEGRATION.md)** | How to use this skill in the project structure             | 10 min    | Project-specific guidance      |
| **[INDEX.md](./INDEX.md)**             | This file - navigation guide                               | 2 min     | Finding what you need          |

### 🎯 Template Files

| File                                                           | Purpose                                        | Usage                                |
| -------------------------------------------------------------- | ---------------------------------------------- | ------------------------------------ |
| **[test-template.py](./test-template.py)**                     | Test file template with fixtures and patterns  | Copy as `tests/test_your_feature.py` |
| **[implementation-template.py](./implementation-template.py)** | Implementation template with endpoint patterns | Copy as `app/routes/your_feature.py` |

---

## 🎯 Common Scenarios

### Scenario 1: Add a new API endpoint

1. **Quick guide**: [CHEATSHEET.md](./CHEATSHEET.md) → "Quick Start" section
2. **Detailed guide**: [INTEGRATION.md](./INTEGRATION.md) → "Adding a New Endpoint"
3. **Real example**: [EXAMPLE.md](./EXAMPLE.md)
4. **Deep dive**: [SKILL.md](./SKILL.md) → "TDD Cycle" sections

**Timeline:**
- 🟢 Beginner: 30 minutes
- 🟡 Intermediate: 15 minutes
- 🔴 Expert: 5 minutes

### Scenario 2: Implement a new utility function

1. **Quick guide**: [CHEATSHEET.md](./CHEATSHEET.md) → "Quick Start" section
2. **Patterns**: [SKILL.md](./SKILL.md) → "Best Practices" section
3. **Templates**: Adapt [test-template.py](./test-template.py) for unit tests
4. **Example**: [EXAMPLE.md](./EXAMPLE.md) → "Phase 3: GREEN" for implementation patterns

### Scenario 3: Fix a bug with confidence

1. **TDD approach**: [SKILL.md](./SKILL.md) → "Test-Driven Development" section
2. **Write test first**: Use [test-template.py](./test-template.py)
3. **Implement fix**: Then write implementation
4. **Verify**: Run full test suite

### Scenario 4: Refactor existing code

1. **Before refactoring**: Ensure tests pass with `pytest tests/ -v`
2. **During refactoring**: Follow [CHEATSHEET.md](./CHEATSHEET.md) → "REFACTOR Phase"
3. **After refactoring**: Run `pytest tests/ -v` again
4. **Commit**: Only commit if all tests still pass

### Scenario 5: Understand project testing patterns

1. **Project structure**: [INTEGRATION.md](./INTEGRATION.md) → "Project Structure"
2. **Endpoint patterns**: [test-template.py](./test-template.py)
3. **Implementation patterns**: [implementation-template.py](./implementation-template.py)
4. **Real examples**: Look at `app/routes/auth.py`, `app/routes/profile.py`

---

## 🚀 Quick Commands

### Run all tests
```bash
pytest tests/ -v
```

### Run specific test file
```bash
pytest tests/test_your_feature.py -v
```

### Run with coverage
```bash
pytest --cov=app tests/
```

### Copy templates
```bash
cp .github/skills/tdd-workflow/test-template.py tests/test_your_feature.py
cp .github/skills/tdd-workflow/implementation-template.py app/routes/your_feature.py
```

### Use skill in Copilot Chat
```
/tdd-workflow [feature description]
```

---

## 📊 TDD Red-Green-Refactor Cycle

```
RED 🔴                  GREEN 🟢               REFACTOR 🔵
Write failing tests  →  Implement minimal code → Improve code quality
Testes fail ❌         Tests pass ✅           Tests still pass ✅
├─ Define behavior      ├─ Make tests pass      ├─ Improve names
├─ Document expected    ├─ Don't over-engineer  ├─ Add docstrings
└─ Create safety net    └─ Focus on passing     └─ Extract helpers
```

**Each phase should result in a commit:**
- `git commit -m "test: add tests for [feature]"`
- `git commit -m "feat: implement [feature]"`
- `git commit -m "refactor: improve [feature] code quality"`

---

## ✨ Key Concepts

### 1. Red-Green-Refactor
The core TDD cycle: write failing tests → implement code → improve quality

### 2. Test-Driven Design
Tests define behavior → guides implementation → creates better architecture

### 3. Regression Prevention
Tests catch when you break existing functionality

### 4. Documentation
Tests show how to use your code → living documentation

### 5. Confidence
Tests make refactoring safe → easier to improve code without fear

---

## 🎓 Learning Path

### For Beginners (New to TDD)
```
1. Read CHEATSHEET.md (3 min)
   ↓
2. Follow EXAMPLE.md step-by-step (15 min)
   ↓
3. Try your first feature with `/tdd-workflow` command
   ↓
4. Read SKILL.md for deeper understanding (15 min)
```

### For Intermediate (Know TDD but new to project)
```
1. Read CHEATSHEET.md (3 min)
   ↓
2. Read INTEGRATION.md for project context (10 min)
   ↓
3. Copy templates and start coding
   ↓
4. Reference SKILL.md as needed
```

### For Advanced (TDD expert)
```
1. Skim CHEATSHEET.md (1 min)
   ↓
2. Copy templates and adapt for your needs
   ↓
3. Reference specific sections as needed
```

---

## 💡 Pro Tips

### Tip 1: Start Small
First feature should be simple (5-10 test cases). Complexity comes with practice.

### Tip 2: Test Behavior, Not Implementation
```python
# ❌ Bad: tests implementation details
def test_function_calls_requests_get():
    with patch("requests.get") as mock:
        ...

# ✅ Good: tests behavior
def test_endpoint_returns_user_data():
    response = client.get("/user")
    assert response.status_code == 200
```

### Tip 3: Mock Everything External
- API calls
- Database queries
- File I/O
- Time/dates (sometimes)

### Tip 4: Commit Frequently
Each phase (red, green, refactor) can be a commit. Small commits = easier to review/revert.

### Tip 5: Run Tests Often
```bash
# While developing
pytest tests/test_your_feature.py -v

# Before committing
pytest tests/ -v

# Before pushing
pytest --cov=app tests/
```

---

## 🔗 External Resources

- [GitHub's TDD Guide](https://github.blog/ai-and-ml/github-copilot/github-for-beginners-test-driven-development-tdd-with-github-copilot/)
- [pytest Documentation](https://docs.pytest.org/)
- [FastAPI Testing](https://fastapi.tiangolo.com/advanced/testing-dependencies/)
- [unittest.mock](https://docs.python.org/3/library/unittest.mock.html)

---

## ❓ FAQ

**Q: Do I have to use TDD for every single thing?**
A: Ideally yes, but at minimum use it for:
- New features
- Bug fixes
- Anything customers depend on

**Q: How many tests do I need?**
A: At least 1 happy path + 1 error case. More coverage = more confidence.

**Q: What if my test is hard to write?**
A: That's a good sign! Hard-to-test code usually has design problems. Rewrite it to be more testable.

**Q: Can I refactor the tests?**
A: Yes, refactor tests too, but ensure they still fail/pass correctly.

**Q: How do I know when to stop testing?**
A: When all edge cases and error paths are covered.

---

## 🎯 Next Steps

Choose your path:

1. **Start immediately**: `/tdd-workflow [your feature]` in Copilot Chat
2. **Learn first**: Read [CHEATSHEET.md](./CHEATSHEET.md) then [EXAMPLE.md](./EXAMPLE.md)
3. **Deep dive**: Read [SKILL.md](./SKILL.md) completely
4. **Integrate**: Read [INTEGRATION.md](./INTEGRATION.md) for project-specific guidance

---

## 📞 Support

If you need help:

1. **Quick questions**: Check [CHEATSHEET.md](./CHEATSHEET.md)
2. **Conceptual questions**: Read [SKILL.md](./SKILL.md)
3. **Project-specific**: Read [INTEGRATION.md](./INTEGRATION.md)
4. **See an example**: Follow [EXAMPLE.md](./EXAMPLE.md)
5. **Get interactive help**: Type `/tdd-workflow [your question]` in Copilot Chat

---

**Happy testing! Remember: Tests written first = code you can trust! 🚀**
