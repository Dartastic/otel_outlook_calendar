// Licensed under the Apache License, Version 2.0
// Copyright 2025-2026 Mindful Software LLC d/b/a Dartastic.io, All rights reserved.

import 'package:dartastic_opentelemetry/dartastic_opentelemetry.dart';
import 'package:otel_outlook_calendar/otel_outlook_calendar.dart';

/// Traces Microsoft Graph calendar calls with OpenTelemetry.
///
/// `otel_outlook_calendar` depends on the OpenTelemetry API only, so the application
/// initializes an SDK once at startup. Wrap each Microsoft Graph calendar call in
/// [tracedOutlookCalendar]; the `invoke` callback makes the real call and reports what
/// came back, which is recorded on the span.
Future<void> main() async {
  await OTel.initialize(serviceName: 'outlook-calendar-example');

  // Wraps a Microsoft Graph calendar call in a CLIENT span.
  final outlookCalendarId = await tracedOutlookCalendar<String>(
    operation: 'events.create',
    mailbox: 'example',
    calendarId: 'calendar_8f3a2c',
    invoke: () async {
      // Make the Microsoft Graph calendar call here, then report what it returned.
      const id = 'outlook_calendar_8f3a2c';
      return const OutlookCalendarCallResult(
          value: id,
          result: OutlookCalendarResult(
              itemCount: 3,
              attendeeCount: 3,
              usedDeltaToken: true,
              deltaTokenReturned: true,
              deltaLinkExpired: true,
              onlineMeetingCreated: true,
              onlineMeetingProvider: 'teamsForBusiness',
              confidence: 'meetingTimeSuggestions[].confidence',
              eventId: 'event_8f3a2c',
              requestId: 'request-id',
              statusCode: 201));
    },
  );
  OTelLog.info(
      'tracedOutlookCalendar events.create returned $outlookCalendarId');

  // Calls that should not be traced, such as health checks, run inside
  // runWithoutOutlookCalendarInstrumentationAsync.
  await runWithoutOutlookCalendarInstrumentationAsync(
    () => tracedOutlookCalendar<String>(
      operation: 'events.create',
      mailbox: 'example',
      calendarId: 'calendar_8f3a2c',
      invoke: () async {
        const id = 'outlook_calendar_8f3a2c';
        return const OutlookCalendarCallResult(
            value: id,
            result: OutlookCalendarResult(
                itemCount: 3,
                attendeeCount: 3,
                usedDeltaToken: true,
                deltaTokenReturned: true,
                deltaLinkExpired: true,
                onlineMeetingCreated: true,
                onlineMeetingProvider: 'teamsForBusiness',
                confidence: 'meetingTimeSuggestions[].confidence',
                eventId: 'event_8f3a2c',
                requestId: 'request-id',
                statusCode: 201));
      },
    ),
  );

  await OTel.shutdown();
}
