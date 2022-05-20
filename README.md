# Inbox Wizard: A Macro & Custom Excel Tables to Automate Inbox Management

## Purpose

At a previous job with an investment advisory firm, managing our team's shared
email inbox was a pain point. We received hundreds of emails each week from
various investment managers and funds in which our clients were invested.
One of our team member's was tasked with reading and processing all of these
emails on a regular basis. I happened to think that her time could be used more
productively, so I set out to relieve her of this mundane task, with the power
of programming!

## Macro Description & Details

Many of these emails included information that we needed for our basic
reporting--namely, monthly statements of our clients' private investments. These
statements provide updates on the current net asset value (NAV) of the investment,
and quantify any additions or withdrawals made during the period.
Without this information, we wouldn't be able to prepare accurate balance sheets
or performance reports for clients. And, unlike your typical stock or mutual fund
investments made in a brokerage account, these private investments are way
behind in terms of reporting, so they did not provide information electronically.
Instead, it was provided in a PDF statement attached to an email.

We also received lots of additional emails that we didn't need to do anything
with at the time, but instead needed to be filed away in the appropriate
inbox subfolder for later reference.

In a given month, we received more than 500 emails, including 200 with these
account statements that needed saving. Imagine sitting
down at your computer on a Friday afternoon, seeing a bold **245 unread** in your
team's shared email inbox, and knowing that you need to open every single one 
of those emails and save an attachment from each (saving to a different
network folder each time). And that your teammates are waiting for you to do this
so they can use the information contained therein.

To relieve the very slow, tedious, and frankly, dehumanizing effort of our team
member manually doing this, I wrote this macro, dubbed the 'Inbox Wizard'.

The macro scans the inbox, searching for recognized senders and emails. If we have
taught it what to do with an email, it will do it...in the blink of an eye.
The macro parses information found in the email subject, body, and attachment names,
and cross-references these data points against the Excel data tables. The
data tables provide a user-friendly way to pass knowledge to the macro. In the
data tables, you can:

<ul>
    <li>add frequent sender email addresses and map them to Fund Managers,</li>
    <li>add keywords to identify emails with account statements vs. emails 
    that only require filing, like quarterly reports, performance estimates, etc.,</li>
    <li>specify the pattern the manager uses to identify our client
    and the specific investment (i.e. 'DATE_CLIENT-ID_FUND-NAME'),</li>
    <li>add client investments, mapping the manager's client ID to our internal
    client ID, allowing the macro to automatically save client statements in the
    correct network folder, and even give the file a friendly name, if we'd like.</li>
</ul>

After the macro runs (either on the entire inbox folder, or the selected emails),
it provides a concise report of its actions, reporting how many emails were evaluated,
how many were skipped as 'unrecognized', and breaks out how many recognized
emails were found *per Manager*, and how many of those were skipped, filed, and/or saved.

Further, to provide tracing and evaluate whether the macro was making any 
unexpected decisions, it keeps logs of all actions taken, saving them to text
files for future reference.

## Conclusion

At the end of the day, I'd estimate that the macro, on a monthly basis, saved
over an hour of (mind-numbing) work. This estimate assumes an average of 250
statements saved per month, and that this would take a person about 10 seconds per
statement, and that it files an additional 250 emails per month, each of which would
take a person 5 seconds to identify and file.

Over the course of a year, that's a **day and a half** back. Worth it to me, since
I mostly did this for fun on my own time (I literally stayed in my room for about 30
hours straight, while my roommates were partying, with a textbook on Visual Basic
that I rented from the library). I guess I should've known at that time that I
was destined to be a software engineer rather than a financial advisor, but it 
would be a full 5 more years before I realized that. Oh well! No regrets.

Thanks for listening.
