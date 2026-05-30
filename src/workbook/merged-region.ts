export class MergedRegion {
  constructor(
    readonly start: string,
    readonly end: string,
  ) {}

  ref(): string {
    return `${this.start}:${this.end}`;
  }
}
