import { Pipe, PipeTransform } from '@angular/core';

@Pipe({
  name: 'currencyOrDash',
  standalone: true
})
export class CurrencyOrDashPipe implements PipeTransform {
  transform(value: number | null | undefined): string {
    if (value === null || value === undefined || isNaN(value)) {
      return '--';
    }

    return value.toLocaleString('pt-BR', {
      style: 'currency',
      currency: 'BRL',
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
  }
}
