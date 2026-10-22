# FX Hedging for a Corporate Exporter: a Sales Case Study (Excel / VBA)

A French company sells to US customers and receives **USD 1m every month for the next
12 months**. Its costs are in euros, so every rise in EURUSD reduces its margin. The treasurer
has a **budget rate of 1.1800** and asks the bank's FX sales desk what to do.

This project puts me in the sales' seat: understand the client's risk, price the hedges the
desk would propose, compare them on the same basis and turn the numbers into a recommendation
the treasurer can act on. Everything is built in one Excel workbook.

![Pricer sheet](docs/pricer.png)

## The client's problem

| | |
|---|---|
| Client | French exporter, receives USD 12m over 12 months |
| Risk | EURUSD goes up: each USD buys fewer euros |
| Budget | 1.1800 (the rate at which the business plan still works) |
| Market | EURUSD 1.1650, 12-month average forward 1.1743 |

Doing nothing is not neutral. In our simulations, the unhedged company misses its budget in
**38% of scenarios**, and in the worst 5% it ends up with an average rate of 1.2549.

## The five solutions a sales would pitch

| Solution | Idea for the client | Cost |
|---|---|---|
| **Forward** | Fix the rate today for every month | Free, but no upside |
| **Vanilla option** | Insurance: protected at 1.1800, keep all the upside | Premium paid upfront |
| **Zero-cost collar** | Protected at 1.1800, upside kept down to a second strike | Free, upside is capped |
| **Participating forward** | Protected at a worse rate, keep 50% of the upside | Free |
| **TARF** | Much better rate, but stops after a target gain and doubles the amount on bad fixings | Free, but risky |

## Results

| Solution | Strike(s) | Premium (EUR) | Expected rate | Worst 5% | Best 5% | Misses budget |
|---|---|---:|---:|---:|---:|---:|
| No hedge | | 0 | 1.1637 | 1.2549 | 1.0756 | 38% |
| Forward | 1.1743 | 0 | 1.1742 | 1.1742 | 1.1742 | 0% |
| Vanilla option | 1.1800 | 197,551 | 1.1678 | 1.2036 | 1.0951 | 49% |
| Zero-cost collar | 1.1800 / 1.1689 | 0 | 1.1736 | 1.1800 | 1.1689 | 0% |
| Participating forward | 1.1926, 50% | 0 | 1.1705 | 1.1922 | 1.1311 | 42% |
| TARF | 1.1245, 1,000 pips, 2x | 0 | 1.1684 | 1.2434 | 1.1073 | 35% |

For an exporter, a **lower rate is better** (more euros for each dollar).

## What I would tell the treasurer

- **The forward is the only solution that guarantees the budget.** Its cost is the 93 pips of
  forward points, which come from USD rates being higher than EUR rates.
- **The option keeps all the upside, but the premium hurts.** EUR 198k paid upfront pushes the
  worst case to 1.2036, above the budget. It suits a client who strongly believes EURUSD will
  fall.
- **The zero-cost collar is the best fit for this client.** No premium, the budget is the worst
  case, and the company still benefits if EURUSD falls to 1.1689.
- **The TARF looks very attractive, but it is a product for experienced clients only.** Its
  strike is almost 500 pips better than the forward. But it stops after 1,000 pips of gains
  (84% of the time, after about 5 months), which leaves the rest of the year unhedged. And if
  EURUSD rises, the company must sell twice its real flows: this happens at least once in 29%
  of scenarios.

My recommendation: hedge most of the flows with the **collar**, and possibly part with
**forwards** to lock in a certain rate. The TARF is shown only for discussion.

The workbook writes this as a one-page client pitch that updates when the inputs change:

![Pitch sheet](docs/pitch.png)

