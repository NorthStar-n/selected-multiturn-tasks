# Dock extract handling

Document: OPS-RUN-DOCK-07
Owner: Eleanor Watts, Logistics Planning
Date: 2022-12-15

- The dock export stores receipt dates as DD/MM/YYYY.
- Status values are copied from the handheld scanner and may contain trailing spaces.
- Negative quantities are quality reversals and reduce the fulfilled quantity.
- Late WMS matcher entries can reference `order_ref` when the `po_number` field is not emitted.
