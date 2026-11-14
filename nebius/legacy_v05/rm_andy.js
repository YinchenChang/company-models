function RM({
  checked: e,
  onChange: t,
  label: n,
  hint: r
}) {
  return (0, $.jsxs)(`label`, {
    className: `flex cursor-pointer items-start gap-2.5 py-1`,
    children: [(0, $.jsx)(`input`, {
      type: `checkbox`,
      checked: e,
      onChange: e => t(e.target.checked),
      className: `mt-1 size-4 accent-accent`
    }), (0, $.jsxs)(`span`, {
      children: [(0, $.jsxs)(`span`, {
        className: `block text-sm font-medium`,
        children: [n, r ? (0, $.jsx)(tipQ, {
          t: r,
          w: 340
        }) : null]
      })]
    })]
  })
}
