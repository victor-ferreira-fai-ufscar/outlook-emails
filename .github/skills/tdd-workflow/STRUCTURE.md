# TDD Workflow Skill - Complete Structure

## 📦 Skill Directory Structure

```
.github/skills/tdd-workflow/
├── 📋 Documentation Files
│   ├── INDEX.md                    # Navigation guide (START HERE!)
│   ├── CHEATSHEET.md               # Quick reference card ⚡ 3 min
│   ├── README.md                   # Skill overview 📖 5 min
│   ├── SKILL.md                    # Complete TDD guide 📚 15 min
│   ├── EXAMPLE.md                  # Step-by-step real example 🎯 10 min
│   ├── INTEGRATION.md              # Project integration guide 🔧 10 min
│   └── STRUCTURE.md                # This file
│
└── 🎯 Template Files
    ├── test-template.py            # Copy → tests/test_[feature].py
    └── implementation-template.py  # Copy → app/routes/[feature].py
```

## 📚 Documentation File Map

```
Choose Your Path Based on Time & Knowledge:

┌─────────────────────────────────────────────────────┐
│         START: INDEX.md (2 min)                     │
│       Navigation guide & quick links                │
└────────────────┬────────────────────────────────────┘
                 │
    ┌────────────┼────────────┬──────────────┐
    │            │            │              │
    ▼            ▼            ▼              ▼
[BUSY]      [LEARNING]   [DEEP DIVE]   [PROJECT]
5 min        10-15 min     30 min        10 min

    │            │            │              │
    ▼            ▼            ▼              ▼
CHEATSHEET → EXAMPLE    → SKILL.md   → INTEGRATION.md
.md        → README.md     .md          .md
Quick       Real          Theory        How-To
Patterns    Example       + Practice     + Patterns

Usage:
1. Copy test template
2. Copy impl template
3. Code with guidance
```

## 🎯 Quick Reference: Where to Look

| Need                | File           | Time   |
| ------------------- | -------------- | ------ |
| **Quick lookup**    | CHEATSHEET.md  | 3 min  |
| **Understand TDD**  | README.md      | 5 min  |
| **Real example**    | EXAMPLE.md     | 10 min |
| **Complete guide**  | SKILL.md       | 15 min |
| **Project details** | INTEGRATION.md | 10 min |
| **File navigation** | INDEX.md       | 2 min  |
| **Visual overview** | STRUCTURE.md   | 2 min  |

## 📋 CHEATSHEET.md Contents

```
├── Quick Start (2 min workflow)
├── TDD Checklist
│   ├── RED Phase
│   ├── GREEN Phase
│   └── REFACTOR Phase
├── Common Test Patterns (6 examples)
├── Common Implementation Patterns (3 examples)
├── pytest Commands
├── Good Practices (5 rules)
├── FAQ Quick (6 Q&A)
└── Workflow Summary Table
```

## 📖 README.md Contents

```
├── Quick Start
├── Why TDD? (5 benefits)
├── TDD Cycle Overview
├── Running Tests
├── Best Practices (5 categories)
├── Project-Specific Notes
├── Resources
└── Need Help?
```

## 📚 SKILL.md Contents

```
├── Overview
├── Why TDD?
├── RED Phase (Write Failing Tests)
│   ├── What to do
│   ├── Checklist
│   └── Example code
├── GREEN Phase (Implement Minimal Code)
│   ├── What to do
│   ├── Checklist
│   └── Example code
├── REFACTOR Phase (Clean Up)
│   ├── What to do
│   ├── Checklist
│   └── Example code
├── Running Tests
├── Best Practices (Detailed)
├── TDD Workflow Checklist
├── Complete Example
├── Resources
└── FAQ with Answers
```

## 🎯 EXAMPLE.md Contents

```
├── Scenario: Add /data/summary endpoint
│
├── Phase 1: RED - Write Failing Tests
│   ├── Step 1: Create test file
│   ├── Step 2: Edit with test code
│   ├── Step 3: Run tests (should FAIL)
│   └── Output showing failure
│
├── Phase 2: GREEN - Implement Minimal Code
│   ├── Step 1: Create implementation file
│   ├── Step 2: Edit with endpoint code
│   ├── Step 3: Register router in main.py
│   ├── Step 4: Run tests (should PASS)
│   └── Output showing success
│
├── Phase 3: REFACTOR - Improve Code Quality
│   ├── Step 1: Refactor implementation
│   ├── Step 2: Run full test suite
│   └── Output showing all still pass
│
├── Phase 4: Commit with Clear History
│   └── 3 commits: test → feat → refactor
│
└── Summary Table showing each phase
```

## 🔧 INTEGRATION.md Contents

```
├── Skill Location
├── When to Use This Skill
├── Using in Copilot Chat (3 methods)
├── Skill Files Overview (table)
├── TDD Workflow in This Project
│   ├── Project Structure
│   └── Adding a New Endpoint (step by step)
├── Running Tests (various commands)
├── Expected Test Coverage
├── Customizing Templates
├── Learning Resources
├── Best Practices (project-specific)
├── Troubleshooting (4 common issues)
└── Quick Navigation
```

## 📋 INDEX.md Contents

```
├── Navigation Guide (choose path)
├── File Directory (table)
├── Common Scenarios (5 examples)
├── Quick Commands
├── TDD Red-Green-Refactor Cycle (visual)
├── Key Concepts (5 ideas)
├── Learning Paths
│   ├── For Beginners
│   ├── For Intermediate
│   └── For Advanced
├── Pro Tips (5 tips)
├── External Resources
├── FAQ (6 Q&A)
└── Next Steps
```

## 🎯 Template Contents: test-template.py

```python
├── Imports
├── Setup Fixtures
│   ├── use_tmp_sessions (auto-cleanup)
│   └── authenticated_session (reusable auth)
├── Happy Path Tests (3 examples)
│   ├── Endpoint exists
│   ├── Returns required fields
│   └── Returns correct data
├── Error Case Tests (2 examples)
│   ├── Requires authentication
│   ├── Rejects invalid session
│   └── Handles API errors
├── Edge Case Tests (2 examples)
│   ├── Empty response
│   └── Special characters (unicode)
└── Instructions (comments)
```

## 🎯 Template Contents: implementation-template.py

```python
├── Imports
├── Router Setup
├── PUBLIC ENDPOINTS
│   ├── Example endpoint
│   └── Full docstring
├── PRIVATE HELPER FUNCTIONS
│   ├── _fetch_external_data()
│   └── _process_data()
├── Advanced Example (commented)
├── Router Registration Instructions
└── Usage Comments
```

## 🔄 Workflow Flow

```
Start Feature Request
        │
        ▼
Use /tdd-workflow command
or read INDEX.md
        │
    ┌───┴────┬──────────┐
    │        │          │
    ▼        ▼          ▼
[RED]   [GREEN]   [REFACTOR]
Write   Implement  Improve
Tests   Code       Quality
    │        │          │
    ├────────┴──────────┤
    │                   │
    ▼                   ▼
pytest FAIL         pytest PASS
    │                   │
    └─────────┬─────────┘
              │
              ▼
        All Tests PASS ✅
              │
              ▼
         Commit Code
              │
              ▼
         Feature Ready! 🚀
```

## 🎓 Learning Timeline

```
Beginner Timeline:
Day 1: Read CHEATSHEET.md (3 min)
Day 1: Follow EXAMPLE.md (15 min)
Day 1: Write first feature with /tdd-workflow
Day 2: Read README.md (5 min)
Day 2: Implement second feature
Day 3: Read SKILL.md (15 min)
Day 3: Implement features with confidence

Intermediate Timeline:
Hour 1: Skim CHEATSHEET.md (3 min)
Hour 1: Read INTEGRATION.md (10 min)
Hour 1: Copy templates and start coding
Hour 2: Reference docs as needed
Day 1+: Implement features using TDD

Advanced Timeline:
Minute 1: Skim CHEATSHEET.md (1 min)
Minute 2: Copy templates
Minute 3: Start coding
Reference: Docs as needed
```

## 📊 Skill Statistics

```
Total Documentation: ~2,000 lines
Code Examples: 50+
Test Patterns: 8
Implementation Patterns: 5
Best Practices: 20+
Learning Paths: 3
Files Created: 8
Quick Start Time: 5 minutes
Complete Understanding: 30 minutes
```

## ✨ Key Features

```
✅ Multiple Entry Points
   ├── Slash command (/tdd-workflow)
   ├── Quick reference (CHEATSHEET.md)
   ├── Real example (EXAMPLE.md)
   └── Complete guide (SKILL.md)

✅ Multiple Learning Styles
   ├── Visual (flowcharts, tables)
   ├── Text (detailed explanations)
   ├── Examples (copy-paste code)
   └── Hands-on (templates)

✅ Project-Specific
   ├── FastAPI patterns
   ├── pytest patterns
   ├── Authentication patterns
   ├── Mock patterns
   └── Router registration patterns

✅ Comprehensive Coverage
   ├── Theory (why TDD works)
   ├── Practice (how to do TDD)
   ├── Patterns (common solutions)
   ├── Troubleshooting (when stuck)
   └── Integration (how to use here)

✅ Quick & Deep
   ├── 5 min quick start
   ├── 15 min example
   ├── 30 min complete
   └── Reference material
```

## 🎯 Success Metrics

After using this skill, you'll be able to:

```
✅ Write failing tests first (RED)
✅ Implement minimal code (GREEN)
✅ Refactor with confidence (REFACTOR)
✅ Use pytest effectively
✅ Mock external APIs
✅ Test authentication patterns
✅ Handle error cases
✅ Test edge cases
✅ Maintain test coverage
✅ Commit with clear history
```

## 🚀 Ready to Use

```
The skill is located in: .github/skills/tdd-workflow/
VS Code Copilot will auto-discover it!

Quick Start:
1. Type: /tdd-workflow [feature description]
2. Follow Copilot's guidance
3. Use templates if needed
4. Reference docs for help

OR

1. Read: CHEATSHEET.md (3 min)
2. Follow: EXAMPLE.md (10 min)
3. Code: Copy templates
4. Reference: Other docs as needed
```

---

**Everything you need to practice TDD is included! 🎉**
