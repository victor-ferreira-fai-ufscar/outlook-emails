<!--- 
How to Access the TDD Workflow Skill
-->

# How to Access the TDD Workflow Skill

## ✨ Method 1: Using the Slash Command (Recommended)

In **Copilot Chat**, type:

```
/tdd-workflow [feature description]
```

### Examples:

```
/tdd-workflow Add a POST endpoint to send emails via Microsoft Graph API

/tdd-workflow Create a utility function to format email addresses

/tdd-workflow Add caching to the /profile endpoint
```

VS Code Copilot will automatically:
1. Load the TDD skill
2. Guide you through the Red phase (write failing tests)
3. Guide you through the Green phase (implement code)
4. Guide you through the Refactor phase (clean up)

---

## 📖 Method 2: Read the Documentation

All skill files are in `.github/skills/tdd-workflow/`:

### Start Here (Based on Your Time)

| Time       | Start With                       | Then Read                          |
| ---------- | -------------------------------- | ---------------------------------- |
| ⚡ 5 min    | [CHEATSHEET.md](./CHEATSHEET.md) | Copy templates, start coding       |
| 📖 15 min   | [EXAMPLE.md](./EXAMPLE.md)       | [INTEGRATION.md](./INTEGRATION.md) |
| 🎓 30 min   | [README.md](./README.md)         | [SKILL.md](./SKILL.md)             |
| 🗺️ Any time | [INDEX.md](./INDEX.md)           | Pick your path                     |

---

## 📋 Method 3: Use the Templates

### Quick Start (Copy & Adapt)

```bash
# 1. Copy test template
cp .github/skills/tdd-workflow/test-template.py tests/test_your_feature.py

# 2. Write your tests in tests/test_your_feature.py
# 3. Run to see them FAIL (RED phase)
pytest tests/test_your_feature.py -v

# 4. Copy implementation template
cp .github/skills/tdd-workflow/implementation-template.py app/routes/your_feature.py

# 5. Write your code to make tests PASS (GREEN phase)
# 6. Register router in app/main.py
# 7. Run tests to see them PASS
pytest tests/test_your_feature.py -v

# 8. Refactor if needed (REFACTOR phase)
# 9. Run full test suite
pytest tests/ -v
```

---

## 🎯 The Quickest Path

If you have **less than 5 minutes**:

1. Type `/tdd-workflow [your feature]` in Copilot Chat
2. Copilot will guide you through each phase
3. Done! 🚀

---

## 📚 The Learning Path

If you want to **understand TDD deeply**:

1. Read [CHEATSHEET.md](./CHEATSHEET.md) (3 min)
2. Follow [EXAMPLE.md](./EXAMPLE.md) step-by-step (10 min)
3. Try your first feature with `/tdd-workflow` command
4. Read [SKILL.md](./SKILL.md) for theory (15 min)

---

## 🔧 For Project-Specific Help

Read [INTEGRATION.md](./INTEGRATION.md) for:
- How to add new endpoints
- Where to put test files
- How to register routers
- Common patterns in this project
- Troubleshooting tips

---

## 📍 Skill Location

```
outlook-emails/
└── .github/skills/tdd-workflow/
    ├── SKILL.md                 ← Main skill file
    ├── README.md
    ├── CHEATSHEET.md
    ├── EXAMPLE.md
    ├── INTEGRATION.md
    ├── INDEX.md
    ├── STRUCTURE.md
    ├── test-template.py
    └── implementation-template.py
```

VS Code Copilot automatically discovers skills in `.github/skills/*/`

---

## ✅ Confirm Skill is Available

In Copilot Chat, type:
```
/tdd-workflow
```

If you see a help message or guidance, the skill is working! ✨

---

## 🆘 Troubleshooting

### "Skill not found" or "/tdd-workflow not recognized"

**Solution:**
1. Restart Copilot Chat (close and reopen)
2. Ensure you're in the correct workspace
3. Check that `.github/skills/tdd-workflow/SKILL.md` exists

### Copilot Chat isn't guiding me through the phases

**Solution:**
1. Type your feature description more clearly
2. Example: `/tdd-workflow Add /emails/send endpoint to send emails`
3. Provide more context about what you want to build

### I prefer to follow the documentation

**Solution:**
1. Open [INDEX.md](./INDEX.md) for navigation
2. Open [CHEATSHEET.md](./CHEATSHEET.md) for quick reference
3. Open [EXAMPLE.md](./EXAMPLE.md) to follow real example
4. Use templates to copy-paste structure

---

## 🎓 Three Ways to Learn

```
┌──────────────────────────────────────────────┐
│ Pick One Approach Below:                     │
└──────────────────────────────────────────────┘

🚀 FASTEST (5 min)
/tdd-workflow [description]
↓
Copilot guides you through phases
↓
Copy templates if needed

📖 BALANCED (20 min)
Read CHEATSHEET.md
↓
Follow EXAMPLE.md
↓
Use templates
↓
Code with Copilot help

🎓 THOROUGH (30+ min)
Read INDEX.md
↓
Read README.md + SKILL.md
↓
Follow EXAMPLE.md
↓
Use INTEGRATION.md for project context
↓
Code with deep understanding
```

---

## 🎯 Next Steps

Choose one:

1. **Try it now**: Type `/tdd-workflow [your feature description]` in Copilot Chat
2. **Learn first**: Read [CHEATSHEET.md](./CHEATSHEET.md)
3. **See an example**: Follow [EXAMPLE.md](./EXAMPLE.md)
4. **Go deep**: Read [SKILL.md](./SKILL.md)

---

**Start implementing features with confidence! 🚀**
