# Comprehensive Markdown Test

This document tests all supported markdown features.

## Text Formatting

This paragraph demonstrates **bold text**, *italic text*, `inline code`, ~~strikethrough~~, and combinations like ***bold italic***.

## Lists

### Unordered List

- First item
- Second item with **bold**
- Third item with *italic*
  - Nested item 1
  - Nested item 2

### Ordered List

1. First numbered item
2. Second numbered item
3. Third numbered item

## Code Blocks

Here's a Python code block:

```python
def fibonacci(n):
    if n <= 1:
        return n
    return fibonacci(n-1) + fibonacci(n-2)

# Calculate first 10 fibonacci numbers
for i in range(10):
    print(f"F({i}) = {fibonacci(i)}")
```

And a JavaScript example:

```javascript
function greet(name) {
    return `Hello, ${name}!`;
}

console.log(greet("World"));
```

## Tables

| Feature | Supported | Notes |
|---------|-----------|-------|
| Headings | Yes | H1 through H6 |
| Bold | Yes | **bold** |
| Italic | Yes | *italic* |
| Tables | Yes | With alignment |
| Images | Yes | Embedded |

## Links

Check out [GitHub](https://github.com) and [Google](https://google.com) for more information.

## Blockquotes

> This is a blockquote.
> It can span multiple lines.
>
> And multiple paragraphs.

## Horizontal Rules

Here's a horizontal rule:

---

And another section below it.

## Mixed Content

You can mix **bold with `code`** and *italic with ~~strikethrough~~* in the same text.

### Task Lists

- [x] Implement markdown parser
- [x] Support tables
- [ ] Add more features
- [ ] Write documentation

## End

This concludes the comprehensive markdown test.
