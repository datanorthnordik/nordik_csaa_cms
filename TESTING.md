# Test Guide

This project includes unit tests (Vitest) and end-to-end tests (Cypress).

## Unit Tests (Vitest)

### Run Tests

- **Watch mode** (development):
  ```bash
  npm test
  # or
  npm run test:unit:watch
  ```

- **Single run**:
  ```bash
  npm run test:unit
  ```

- **With coverage**:
  ```bash
  npm run test:unit:coverage
  ```

### Test Files

- Unit tests should be named `*.test.ts` or `*.test.tsx`
- Spec tests should be named `*.spec.ts` or `*.spec.tsx`
- Test files are colocated with the source code they test

## End-to-End Tests (Cypress)

### Run Tests

- **Interactive mode** (Cypress UI):
  ```bash
  npm run test:e2e:open
  ```

- **Headless mode**:
  ```bash
  npm run test:e2e:run
  ```

- **Against dev server**:
  ```bash
  npm run test:e2e
  ```

### Test Files

- E2E tests are located in `cypress/e2e/`
- Tests should be named `*.cy.ts` or `*.cy.tsx`

## Combined Tests

### Run All Tests

- **Unit + E2E tests** (requires built app):
  ```bash
  npm run test:ci
  ```

### Quality Checks

- **Linting + unit tests** (for PRs/pre-commit):
  ```bash
  npm run test:check
  ```

## Coverage

Unit test coverage is generated in the `coverage/` directory. Coverage thresholds are set to 70% for:
- Lines
- Functions
- Branches
- Statements

View HTML report:
```bash
open coverage/index.html
```

## Debugging Tests

### Vitest

Add a debugger statement in your test and run:
```bash
node --inspect-brk ./node_modules/.bin/vitest --inspect-brk --run
```

### Cypress

Use the Cypress UI (`test:e2e:open`) to debug tests interactively.

## Best Practices

1. Keep tests close to the code they test
2. Test behavior, not implementation details
3. Use descriptive test names
4. Mock external dependencies
5. Keep tests isolated and independent
6. Use `beforeEach`/`afterEach` for setup/teardown
