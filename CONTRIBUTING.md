# Contributing

## Setup

```bash
pnpm install
```

## Development

```bash
# Build all packages
pnpm build

# Run tests
pnpm test

# Type-check
pnpm typecheck

# Lint
pnpm lint
pnpm lint:fix

# Format
pnpm format
pnpm format:fix
```

## Project Structure

```
quadro/
├── packages/core/        # @qquadro/core — publishable library
├── packages/tsconfig/    # @quadro/tsconfig — shared TypeScript configs
└── apps/cli/             # @quadro/cli — CLI tool
```

## Pull Requests

1. Ensure all checks pass: `pnpm turbo run build typecheck test lint`
2. Write tests for new functionality
3. Update the changelog
