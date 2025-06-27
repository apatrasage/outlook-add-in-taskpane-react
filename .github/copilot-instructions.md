# Copilot Instructions: Best Practices for This Project

## TypeScript & JavaScript Best Practices

- Always use strict typing in TypeScript; avoid `any` unless absolutely necessary.
- Prefer interfaces over types for object shapes in TypeScript.
- Use enums for fixed sets of values.
- Leverage type inference but annotate function signatures.
- Handle all possible cases in switch statements and discriminated unions.
- Use `const` and `let` appropriately; avoid `var`.
- Prefer arrow functions for callbacks and functional components.
- Avoid mutating objects/arrays directly; use spread/rest operators.
- Use template literals for string concatenation.
- Write clean, readable, and maintainable code.
- Use meaningful variable, function, and component names.
- Add comments and JSDoc/TSDoc where necessary.
- Avoid code duplication; use reusable functions/components.
- Ensure all code is covered by automated tests.

## Office Add-in (Outlook) Best Practices

- Follow [Microsoft Office Add-in Design Guidelines](https://learn.microsoft.com/en-us/office/dev/add-ins/design/add-in-design-guidelines).
- Use Office.js APIs according to [Outlook Add-in documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/).
- Always check for API support using `Office.context.requirements.isSetSupported` before using advanced features.
- Handle errors gracefully and provide user-friendly error messages.
- Never block the UI thread; use async/await or Promises for all Office.js calls.
- Respect user privacy: never store or transmit user data without consent.
- Do not request more permissions than necessary in the manifest.
- Test add-in behavior in all supported Outlook clients (desktop, web, mobile).
- Ensure add-in is resilient to slow or unavailable Office APIs.
- Use localization for all user-facing strings.

## Security & Privacy for Add-ins

- Sanitize all data from Office APIs before rendering or processing.
- Never expose sensitive information in logs or error messages.
- Use HTTPS for all network requests.
- Validate and escape all data used in the DOM to prevent XSS.
- Follow [Office Add-in security best practices](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/security).

## Fluent UI Best Practices

- Use Fluent UI React components for consistent, accessible, and modern UI.
- Prefer Fluent UI's built-in accessibility features and follow their [accessibility documentation](https://react.fluentui.dev/?path=/docs/concepts-accessibility--page).
- Use `makeStyles` or Fluent UI's styling solutions instead of inline styles.
- Leverage Fluent UI tokens and theme variables for color, spacing, and typography.
- Use semantic components (e.g., `Button`, `Label`, `Card`) and provide appropriate ARIA attributes.
- Ensure all custom components built on Fluent UI primitives maintain keyboard accessibility and focus management.
- Use `mergeClasses` for combining style classes.
- Avoid overriding Fluent UI styles unless necessary; extend with custom styles using recommended APIs.
- Test components with screen readers and keyboard navigation to ensure accessibility.
- Reference: [Fluent UI React Documentation](https://react.fluentui.dev/)

## React Best Practices

- Use functional components and React hooks.
- Keep components small and focused.
- Use prop-types or TypeScript interfaces for props.
- Avoid inline styles; use CSS-in-JS or CSS modules.
- Use keys for list rendering.
- Avoid side effects in render; use `useEffect` for side effects.

## Accessibility (EU Accessibility Act, WCAG 2.1 AA, Office Add-in)

- All interactive elements must be keyboard accessible (tab, enter, space).
- Use semantic HTML elements (e.g., `<button>`, `<nav>`, `<main>`, `<header>`, `<footer>`).
- Provide `aria-label`, `aria-labelledby`, or `aria-describedby` where necessary.
- Ensure sufficient color contrast (minimum 4.5:1 for normal text).
- All images must have descriptive `alt` text.
- Use focus indicators for all focusable elements.
- Avoid using color as the only means of conveying information.
- Support screen readers: use ARIA roles and properties appropriately.
- Test with keyboard-only navigation and screen readers.
- Avoid auto-playing media or flashing content.
- Ensure all forms have associated labels and error messages are accessible.
- Follow [Office Add-in accessibility guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/design/accessibility-checklist).
- Ensure all add-in commands and UI are accessible via keyboard and screen reader in Outlook clients.
- Test with Narrator, NVDA, or VoiceOver in addition to browser tools.
- Reference: [EU Accessibility Act](https://commission.europa.eu/strategy-and-policy/policies/justice-and-fundamental-rights/disability/union-equality-strategy-rights-persons-)

## Testing

- Write unit tests for all functions, components, and utilities.
- Write integration tests for component interactions and API calls.
- Cover all edge cases and error states.
- Use mocks/stubs for external dependencies.
- Test accessibility: use tools like axe, jest-axe, or @testing-library/jest-dom.
- Ensure 100% test coverage for all new code.
- Add regression tests for fixed bugs.
- Test add-in in all supported Outlook environments (Windows, Mac, Web, Mobile).
- Use [Office Add-in Validator](https://github.com/OfficeDev/office-addin-validator) and [Accessibility Insights](https://accessibilityinsights.io/) for compliance.

## Example Test Cases

- Render components with all possible prop combinations.
- Test keyboard navigation and focus management.
- Test ARIA attributes and screen reader output.
- Test error boundaries and fallback UI.
- Test API error and success responses.
- Test form validation and error messages.
- Test Office.js API error handling and fallback UI.
- Test add-in manifest for permission and requirement set compliance.
- Test localization in all supported languages.

## Pull Request Checklist

- [ ] Code follows best practices for TypeScript, JavaScript, and React.
- [ ] All code is covered by tests (unit, integration, accessibility).
- [ ] All UI is accessible and compliant with the EU Accessibility Act.
- [ ] No accessibility violations (checked with automated tools and manual testing).
- [ ] All tests pass.
- [ ] Add-in manifest follows least-privilege and Office Store requirements.
- [ ] Add-in tested in all supported Outlook clients.

---

_Please follow these guidelines to ensure code quality, maintainability, accessibility, and Office Add-in compliance._
