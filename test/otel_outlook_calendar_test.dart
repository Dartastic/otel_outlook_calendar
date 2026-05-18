// Licensed under the Dartastic Pro Software License Agreement.
// Copyright 2025-2026 Mindful Software LLC d/b/a Dartastic.io, All rights reserved.

import 'package:dartastic_opentelemetry/dartastic_opentelemetry.dart';
import 'package:dartastic_opentelemetry/testing.dart';
import 'package:otel_outlook_calendar/otel_outlook_calendar.dart';
import 'package:test/test.dart';

void main() {
  late TestHarness harness;
  late InMemorySpanExporter spans;

  setUpAll(() async {
    harness = await maybeInitializeOtelForTest(
      serviceName: 'otel_outlook_calendar-test',
    );
    spans = harness.spans;
  });

  setUp(() => harness.clear());

  Map<String, Object> attrsOf(Span s) =>
      {for (final a in s.attributes.toList()) a.key: a.value};

  test('delta sync records used + returned', () async {
    await tracedOutlookCalendar<String>(
      operation: 'delta',
      mailbox: 'user@example.com',
      invoke: () async => const OutlookCalendarCallResult(
        value: 'ok',
        result: OutlookCalendarResult(
          itemCount: 12,
          usedDeltaToken: true,
          deltaTokenReturned: true,
          statusCode: 200,
        ),
      ),
    );
    final a = attrsOf(spans.findSpanByName('outlook_calendar delta')!);
    expect(a['outlook_calendar.mailbox'], 'user@example.com');
    expect(a['outlook_calendar.delta_token.used'], true);
    expect(a['outlook_calendar.delta_token.returned'], true);
  });

  test('410 Gone delta expiry maps to span error', () async {
    await tracedOutlookCalendar<String>(
      operation: 'delta',
      invoke: () async => const OutlookCalendarCallResult(
        value: 'ok',
        result: OutlookCalendarResult(
          deltaLinkExpired: true,
          statusCode: 410,
        ),
      ),
    );
    final s = spans.findSpanByName('outlook_calendar delta')!;
    expect(s.status, SpanStatusCode.Error);
    final a = attrsOf(s);
    expect(a['outlook_calendar.delta_link.expired'], true);
  });

  test('events.create with Teams meeting', () async {
    await tracedOutlookCalendar<String>(
      operation: 'events.create',
      invoke: () async => const OutlookCalendarCallResult(
        value: 'ok',
        result: OutlookCalendarResult(
          attendeeCount: 8,
          eventId: 'AAMk-1',
          onlineMeetingCreated: true,
          onlineMeetingProvider: 'teamsForBusiness',
          statusCode: 201,
        ),
      ),
    );
    final a = attrsOf(spans.findSpanByName('outlook_calendar events.create')!);
    expect(a['outlook_calendar.attendee_count'], 8);
    expect(a['outlook_calendar.online_meeting.created'], true);
    expect(a['outlook_calendar.online_meeting.provider'], 'teamsForBusiness');
  });

  test('suppression skips spans', () async {
    await runWithoutOutlookCalendarInstrumentationAsync(() async {
      await tracedOutlookCalendar<String>(
        operation: 'events.list',
        invoke: () async => const OutlookCalendarCallResult(value: 'v'),
      );
    });
    expect(spans.findSpansStartingWith('outlook_calendar '), isEmpty);
  });
}
