export class CellComment {
  constructor(
    readonly text: string,
    readonly author: string | null = null,
    readonly color: string | null = null,
  ) {}
}
