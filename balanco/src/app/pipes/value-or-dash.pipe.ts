import { Pipe, PipeTransform } from '@angular/core';

@Pipe({
  name: 'valueOrDash',
  standalone: true,
})
export class ValueOrDashPipe implements PipeTransform {
  transform(value: any): any {
    return value === null || value === undefined || value === '' ? '---' : value;
  }
}
