# Inbox Wizard: A Macro to Automate Outlook Inbox Management

## Purpose

At a previous job with an investment advisory firm, managing our team's shared
email inbox was a pain point. We received hundreds of emails each week from
various investment managers and funds in which our clients were invested.
One of our team member's was tasked with reading and processing all of these
emails on a regular basis. I happened to think that her time could be used more
productively, so I set out to relieve her of this mundane task, with the power
of programming!

Above I have included the macro (run from Outlook) and the excel data tables.
All client and other sensitive information has been scrubbed.

## Further Background

Many of these emails included information that we needed for our basic
reporting; namely, monthly statements of our clients' private investments. These
statements provide updates on the current net asset value (NAV) of the investment,
and quantify any additions or withdrawals made during the period.
Without this information, we wouldn't be able to prepare accurate balance sheets
or performance reports for clients. And, unlike your typical stock or mutual fund
investments made in a brokerage account, these private investment managers (or
really their fund administrators) are way
behind in terms of reporting, so they did not provide information electronically.
Instead, it was provided in a PDF statement attached to an email.

We also received lots of additional emails that we didn't include statements we
needed to save, but did need to be filed away in the appropriate
inbox subfolder for later reference. This included email updates like quarterly
reports, estimated performance, timing expectations for tax documents, and various
other updates managers like to provide.

In a given month, we received more than 500 emails, including 200 with these
account statements that needed saving. Imagine sitting
down at your computer on a Friday afternoon, seeing **245 unread** emails in your
team's shared email inbox, and knowing that you need to open every single one 
of those emails and save an attachment from each (saving to a different
network folder each time). And that your teammates are waiting for you to do this
so they can use the information they contain. It makes me sweat just thinking about it.

To relieve the very slow, tedious, and frankly, dehumanizing effort of our team
member doing this manually, I wrote this macro, dubbed the 'Inbox Wizard'.

## How It Works

The macro scans the inbox, searching for recognized senders and emails. If we have
taught it what to do with an email, it will do it...in the blink of an eye.
The macro parses information found in the email subject, body, and attachment names,
and cross-references these data points against the Excel data tables to lookup what
actions to take: a) file the email direclty into a subfolder, b) save an attachment
and then file the email, or c) leave it in the inbox for a human to look at.
The data tables provide a user-friendly way to pass knowledge to the macro. In the
data tables, you can:

<ul>
    <li>add frequent sender email addresses and map them to Fund Managers,</li>
    <li>add keywords to identify emails with account statements vs. emails 
    that only require filing, like quarterly reports, performance estimates, etc.,
    and where to search for those keywords (subject line, body, or attachments),</li>
    <li>specify the pattern the manager uses to identify our client
    and the specific investment (e.g. 'date_client-ID_fund-ID'),</li>
    <li>add client information, mapping the manager's client ID to our internal
    client ID, allowing the macro to automatically save client statements in the
    correct network folder, and even give the file a friendly name, if we'd like.</li>
</ul>

After the macro runs (either on the entire inbox folder, or the selected emails),
it provides a concise report of its actions, reporting how many emails were evaluated,
how many were skipped as 'unrecognized', and breaks out how many recognized
emails were found *per Manager*, and how many of those were skipped, filed, and/or saved.

Further, to help identify more ways the macro could help, and to provide tracing
and evaluate whether the macro was making any unexpected decisions, it keeps
descriptive logs of all actions taken, saving the log files for future reference.

## Conclusion

At the end of the day, I'd estimate that the macro, on a monthly basis, saved
over an hour of (mind-numbing) work. This estimate assumes an average of 250
statements saved per month, and that this would take a person about 10 seconds per
statement (with no breaks), and that it files an additional 250 emails per month, each of which would
take a person 5 seconds to identify and file.

Over the course of a year, that's a **day and a half** back. That may not seem like
much, but it was in use for more than 4 years, and was well worth it to me.
I mostly did this for fun on my own time (I literally stayed in my room for about 30
hours straight, while my roommates were partying, with a textbook on Visual Basic
that I rented from the Chicago Public Library). I guess I should've known at that time that I
was destined to be a software engineer rather than a financial advisor, but it 
would be a full 5 more years before I realized that...oh well! No regerts.

If you take a look at the code, keep in mind that this was really my first large
programming project! And I was not a CS major. So go easy on me. I would've done
some things pretty differently today, but looking back, I'm proud I came up with
the code that I did!

Thanks for reading.
