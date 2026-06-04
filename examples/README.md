# Examples

This folder contains tiny Office OOXML fixtures that show the cleanup workflow.

## Files

- `before/`: intentionally messy examples with Chinese-English spaces, Word/Excel fill colors, and non-normalized fonts.
- `after/`: the same files after running `scripts/enforce_office_fonts.py`.

## Reproduce

From the repository root:

```bash
python3 scripts/enforce_office_fonts.py --recursive examples/before
mkdir -p examples/after
mv examples/before/*_fontfixed.* examples/after/
```

Or check the messy examples:

```bash
python3 scripts/enforce_office_fonts.py --check --recursive examples/before
```

The `before` files should fail check mode. The `after` files should pass.
